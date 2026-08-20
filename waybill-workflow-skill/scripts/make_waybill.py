# -*- coding: utf-8 -*-
"""运单副本生成器：基于旧运单 PDF 模板生成新运单副本（原位编辑）。"""
import argparse, json, os, re, sys, statistics

import fitz

C39 = {
    '0': '000110100', '1': '100100001', '2': '001100001', '3': '101100000', '4': '000110001',
    '5': '100110000', '6': '001110000', '7': '000100101', '8': '100100100', '9': '001100100',
    'A': '100001001', 'B': '001001001', 'C': '101001000', 'D': '000011001', 'E': '100011000',
    'F': '001011000', 'G': '000001101', 'H': '100001100', 'I': '001001100', 'J': '000011100',
    'K': '100000011', 'L': '001000011', 'M': '101000010', 'N': '000010011', 'O': '100010010',
    'P': '001010010', 'Q': '000000111', 'R': '100000110', 'S': '001000110', 'T': '000010110',
    'U': '110000001', 'V': '011000001', 'W': '111000000', 'X': '010010001', 'Y': '110010000',
    'Z': '011010000', '-': '010000101', '.': '110000100', ' ': '011000100', '*': '010010100',
    '$': '010101000', '/': '010100010', '+': '010001010', '%': '000101010'}
C39_CHARS = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%'


def c39_check(data):
    return C39_CHARS[sum(C39_CHARS.index(c) for c in data) % 43]


def c39_elements(text):
    elems = []
    for i, ch in enumerate(text):
        for j, bit in enumerate(C39[ch]):
            elems.append(('bar' if j % 2 == 0 else 'space', 3 if bit == '1' else 1))
        if i < len(text) - 1:
            elems.append(('space', 1))
    return elems


def decode_c39(bars):
    mod = lambda w: 1 if w < 1.8 else 3
    seq = []
    for i, (x0, x1) in enumerate(bars):
        seq.append(mod(x1 - x0))
        if i < len(bars) - 1:
            seq.append(mod(bars[i + 1][0] - x1))
    chars, i = [], 0
    while i + 9 <= len(seq):
        bits = ''.join('1' if w == 3 else '0' for w in seq[i:i + 9])
        ch = [c for c, p in C39.items() if p == bits]
        chars.append(ch[0] if ch else '?')
        i += 10
    return ''.join(chars)


def find_numbers(page):
    """返回 [(rect, text)]：形如 NNN-NNNNNNNN 的大号粗体文本。"""
    out = []
    for b in page.get_text('dict')['blocks']:
        for l in b.get('lines', []):
            for s in l['spans']:
                t = s['text'].strip()
                if re.fullmatch(r'\d{3}-\d{8}', t) and s['size'] >= 12:
                    out.append((fitz.Rect(*s['bbox']), t))
    return out


def find_bars(page):
    cands = []
    for d in page.get_drawings():
        for it in d['items']:
            if it[0] == 're':
                r = it[1]
                if 2 < r.y0 < 20 and (r.y1 - r.y0) > 8 and (r.x1 - r.x0) < 5:
                    cands.append((r.x0, r.x1, r.y0, r.y1))
    cands.sort()
    if not cands:
        return []
    # 按 x 间距聚类，取最大簇（排除其他零散填充矩形）
    clusters, cur = [], [cands[0]]
    for c in cands[1:]:
        if c[0] - cur[-1][1] < 12:
            cur.append(c)
        else:
            clusters.append(cur)
            cur = [c]
    clusters.append(cur)
    bars = max(clusters, key=len)
    return bars if len(bars) >= 40 else []


def find_goods_area(page):
    """右半页小字号连续行段，返回 (rect, fontname, fontsize, line_h, from_text)。
    优先从已有文本检测；若模板无品名文本，则从背景图片中检测表格线定位。"""
    spans = []
    for b in page.get_text('dict')['blocks']:
        for l in b.get('lines', []):
            for s in l['spans']:
                if s['size'] <= 8 and s['bbox'][0] > page.rect.width * 0.5:
                    spans.append(s)
    spans.sort(key=lambda s: s['bbox'][1])
    
    # 尝试从文本检测品名区
    if spans:
        lines = []
        for s in spans:
            if lines and abs(s['bbox'][1] - lines[-1][-1]['bbox'][1]) < 2:
                lines[-1].append(s)
            else:
                lines.append([s])
        segs, cur = [], [lines[0]]
        for ln in lines[1:]:
            if ln[0]['bbox'][1] - cur[-1][0]['bbox'][1] < 8:
                cur.append(ln)
            else:
                segs.append(cur)
                cur = [ln]
        segs.append(cur)
        seg = max(segs, key=len)
        # 只有多行段才是品名区（排除单行的ID文本等）
        if len(seg) >= 3:
            flat = [s for ln in seg for s in ln]
            x0 = min(s['bbox'][0] for s in flat)
            x1 = max(s['bbox'][2] for s in flat)
            y0 = min(s['bbox'][1] for s in flat)
            y1 = max(s['bbox'][3] for s in flat)
            fs = statistics.median(s['size'] for s in flat)
            lhs = [seg[i + 1][0]['bbox'][1] - seg[i][0]['bbox'][1] for i in range(len(seg) - 1)]
            lh = statistics.median(lhs) if lhs else fs * 1.2
            # 检测图片中的右边界竖线
            right_border = _find_right_border_from_image(page, y0, y1)
            rect = fitz.Rect(x0 - 4, y0 - 2, min(x1 + 4, right_border), y1 + 2)
            main = max(flat, key=lambda s: len(s['text']))
            return rect, main['font'], fs, lh, True
    
    # 回退：模板无品名文本，从图片中检测表格线定位品名区
    return _find_goods_area_from_image(page)


def _find_right_border_from_image(page, y0_pt, y1_pt):
    """从页面背景图片中检测品名栏右侧竖线位置。"""
    try:
        pix = page.get_pixmap()
        import numpy as np
        arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape((pix.height, pix.width, pix.n))
        scale_x = pix.width / page.rect.width
        scale_y = pix.height / page.rect.height
        y_start_px = max(0, int(y0_pt * scale_y))
        y_end_px = min(pix.height, int((y1_pt + 10) * scale_y))
        if y_end_px <= y_start_px + 10:
            return page.rect.width - 10
        # 在右侧 85% 以后找暗色列
        threshold = 80
        right_start = int(page.rect.width * 0.85 * scale_x)
        for x in range(right_start, pix.width):
            col_slice = arr[y_start_px:y_end_px, x]
            if col_slice.mean() < threshold:
                return x / scale_x - 2
    except Exception:
        pass
    return page.rect.width - 10


def _find_goods_area_from_image(page):
    """模板无品名文本时，从背景图片检测表格线确定品名区位置。
    返回 (rect, fontname, fontsize, line_h, from_text)。"""
    W, H = page.rect.width, page.rect.height
    # 默认值（基于常见托书布局）
    x0, x1 = W * 0.42, W - 15
    y0, y1 = H * 0.43, H * 0.58
    # 字号对齐已填好的真实托书（英文 4.2/中文 4.0，行距约 5）
    fs, lh = 4.2, 5.0
    fontname = 'Helvetica'
    
    try:
        pix = page.get_pixmap()
        import numpy as np
        arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape((pix.height, pix.width, pix.n))
        scale_x = pix.width / W
        scale_y = pix.height / H
        
        # 找竖线：在页面中部区域(y 40%-60%)找暗色列
        y_mid_start = int(H * 0.35 * scale_y)
        y_mid_end = int(H * 0.65 * scale_y)
        vline_xs = []
        for x in range(pix.width):
            col_slice = arr[y_mid_start:y_mid_end, x]
            if col_slice.mean() < 80:
                vline_xs.append(x)
        # 聚类找竖线
        if vline_xs:
            clusters = []
            cur = [vline_xs[0]]
            for x in vline_xs[1:]:
                if x - cur[-1] <= 3:
                    cur.append(x)
                else:
                    clusters.append(cur)
                    cur = [x]
            clusters.append(cur)
            # 找品名栏左边界（页面中部偏右的竖线）和右边界（最右侧竖线）
            centers = [(min(c) + max(c)) / 2 for c in clusters]
            # 右边界：最右侧的竖线
            right_x = max(centers) / scale_x - 2
            # 左边界：在页面 35%-50% 范围内的竖线
            left_candidates = [c for c in centers if W * 0.30 * scale_x < c < W * 0.55 * scale_x]
            if left_candidates:
                left_x = max(left_candidates) / scale_x + 3
            else:
                left_x = W * 0.42
            x0, x1 = left_x, right_x
        
        # 找横线：在品名栏 x 范围内找水平线确定上下边界
        x_start_px = int(x0 * scale_x)
        x_end_px = int(x1 * scale_x)
        hline_ys = []
        for y in range(int(H * 0.3 * scale_y), int(H * 0.7 * scale_y)):
            row_slice = arr[y, x_start_px:x_end_px]
            if row_slice.mean() < 80:
                hline_ys.append(y)
        if hline_ys:
            # 聚类找横线
            hclusters = []
            cur = [hline_ys[0]]
            for y in hline_ys[1:]:
                if y - cur[-1] <= 3:
                    cur.append(y)
                else:
                    hclusters.append(cur)
                    cur = [y]
            hclusters.append(cur)
            hcenters = [(min(c) + max(c)) / 2 for c in hclusters]
            # 品名区上下边界：找中间区域的横线对
            mid_ys = [y for y in hcenters if H * 0.3 * scale_y < y < H * 0.7 * scale_y]
            if len(mid_ys) >= 2:
                # 第一、二条横线之间是表头行（Nature & Quantity of Goods），
                # 品名内容必须从第二条横线（表头下边框）之后开始，否则会压住表头
                top_idx = 1 if len(mid_ys) >= 3 else 0
                y0 = mid_ys[top_idx] / scale_y + 2
                # 品名格下边界 = 上边界后的下一条横线（而非最后一条）；
                # 最后一条是下方郑重声明区的底边，用它会压住声明文字
                bot_idx = top_idx + 1 if top_idx + 1 < len(mid_ys) else len(mid_ys) - 1
                y1 = mid_ys[bot_idx] / scale_y - 2
    except Exception:
        pass
    
    rect = fitz.Rect(x0, y0, x1, y1)
    return rect, fontname, fs, lh, False


def load_goods_font(doc, page, fontname):
    for f in page.get_fonts(full=True):
        xref, name = f[0], f[3]
        if fontname and (fontname in name or name.endswith(fontname) or fontname.split('+')[-1] in name):
            try:
                nm, ext, ftype, content = doc.extract_font(xref)
                if content:
                    p = os.path.join(os.environ.get('TEMP', '.'), 'waybill_song.' + (ext if ext in ('otf', 'ttf') else 'otf'))
                    with open(p, 'wb') as fh:
                        fh.write(content)
                    return fitz.Font(fontfile=p)
            except Exception:
                pass
    return fitz.Font(fontfile=r"C:\Windows\Fonts\simsun.ttc")


def wrap(text, font, fs, maxw):
    """将文本按 maxw 宽度分行。优先在空格/标点处断开，避免单行溢出。
    对超长英文串，会在逗号、括号等标点处主动换行，而非强制拆字符。"""
    # 分词：保留空格和标点作为独立 token
    # 关键：尾部标点（逗号、句号、右括号等）必须单独成 token，才能作为换行点
    units = re.findall(r'[A-Za-z0-9][A-Za-z0-9.:+/\-]*(?:[A-Za-z0-9.:+/\-]*[A-Za-z0-9])?|[,:;.()\[\]{}]+|\s|.', text)
    lines, cur = [], ''

    def force_split(s):
        """最后手段：逐字符拆分超长串。"""
        out = []
        while font.text_length(s, fontsize=fs) > maxw:
            k = len(s)
            while k > 1 and font.text_length(s[:k], fontsize=fs) > maxw:
                k -= 1
            out.append(s[:k])
            s = s[k:]
        return out, s

    for u in units:
        t = cur + u
        if font.text_length(t, fontsize=fs) > maxw:
            # 当前行已满，先保存已有内容
            if cur.strip():
                lines.append(cur.rstrip())
            # 新单元本身是否超宽？
            if font.text_length(u, fontsize=fs) > maxw:
                # 尝试在标点处拆分（比逐字符拆更自然）
                punct_splits = re.split(r'(?<=[,;.:()\[\]{}])', u)
                if len(punct_splits) > 1:
                    # 有多个片段，逐个放入
                    sub_cur = ''
                    for seg in punct_splits:
                        if not seg:
                            continue
                        test = sub_cur + seg
                        if font.text_length(test, fontsize=fs) <= maxw:
                            sub_cur = test
                        else:
                            if sub_cur.strip():
                                lines.append(sub_cur.rstrip())
                            sub_cur = seg
                            if font.text_length(sub_cur, fontsize=fs) > maxw:
                                extra, sub_cur = force_split(sub_cur)
                                lines.extend(extra)
                    cur = sub_cur
                else:
                    # 没有标点可拆，只能逐字符拆
                    extra, cur = force_split(u)
                    lines.extend(extra)
            else:
                # 单元不超宽，直接作为新行开头
                cur = u if u != ' ' else ''
        else:
            cur = t
    if cur.strip():
        lines.append(cur.rstrip())
    return lines


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--template', required=True)
    ap.add_argument('--new', required=True)
    ap.add_argument('--old', default=None)
    ap.add_argument('--excel', default=None)
    ap.add_argument('--items-json', default=None)
    ap.add_argument('--en-head', default='')
    ap.add_argument('--en-tail', default='')
    ap.add_argument('--cn-tail', default='')
    ap.add_argument('--out', required=True)
    ap.add_argument('--no-ai-label', action='store_true')
    a = ap.parse_args()

    doc = fitz.open(a.template)
    assert len(doc) >= 1
    page = doc[0]
    W = page.rect.width

    # ---- 品名数据 ----
    items = []
    if a.excel:
        import openpyxl
        wb = openpyxl.load_workbook(a.excel)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        hdr = list(rows[0])
        ci_cn = next(i for i, h in enumerate(hdr) if h and '中文' in str(h))
        ci_en = next(i for i, h in enumerate(hdr) if h and '英文' in str(h))
        items = [(r[ci_cn], r[ci_en]) for r in rows[1:] if r[ci_cn]]
    elif a.items_json:
        items = [tuple(x) for x in json.load(open(a.items_json, encoding='utf-8'))]
    en_text = a.en_head + ",".join(en.rstrip(',').rstrip() for _, en in items) + (("." + a.en_tail) if a.en_tail else "")
    cn_text = ",".join(cn.rstrip(',').rstrip() for cn, _ in items) + a.cn_tail

    # ---- 旧号 / 新号 ----
    nums = find_numbers(page)
    if not nums:
        sys.exit('未找到模板运单号')
    old_no = a.old or nums[0][1]
    num_rects = [r for r, t in nums if t == old_no]
    if not num_rects:
        sys.exit(f'模板中未找到旧运单号 {old_no}')

    # ---- 条码 ----
    bars = find_bars(page)
    narrow = statistics.median([b[1] - b[0] for b in bars if (b[1] - b[0]) < 1.8]) if bars else 0.875
    do_barcode = bool(bars) and all(c in C39 for c in '*' + a.new + '0' + '*')
    if bars:
        dec = decode_c39([(b[0], b[1]) for b in bars])
        print('old barcode decoded:', dec)
        if a.old and dec and '?' not in dec and old_no not in dec:
            print('WARN: 条码解码与旧号不符，仍按新号重生成')

    # ---- 品名区 ----
    g_rect, g_fontname, fs, lh, from_text = find_goods_area(page)
    font = load_goods_font(doc, page, g_fontname)

    # ---- 擦除 ----
    rects = list(num_rects) + [g_rect]
    for r in rects:
        page.add_redact_annot(r, fill=None)
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_NONE)
    if bars:
        by0 = min(b[2] for b in bars) - 0.6
        by1 = max(b[3] for b in bars) + 0.6
        page.draw_rect(fitz.Rect(bars[0][0] - 1, by0, bars[-1][1] + 1, by1), color=None, fill=(1, 1, 1))

    # ---- 重画运单号 ----
    for r in num_rects:
        page.insert_text((r.x0, r.y0 + (r.y1 - r.y0) * 0.82), a.new, fontname="hebo",
                         fontsize=14.0, color=(0, 0, 0))

    # ---- 重画条码 ----
    if bars and do_barcode:
        full = "*" + a.new + c39_check(a.new) + "*"
        x = bars[0][0]
        shape = page.new_shape()
        for kind, w in c39_elements(full):
            width = w * narrow
            if kind == 'bar':
                shape.draw_rect(fitz.Rect(x, bars[0][2], x + width, bars[0][3]))
            x += width
        shape.finish(color=None, fill=(0, 0, 0))
        shape.commit()
        print('new barcode width:', round(x - bars[0][0], 2))

    # ---- 重画品名 ----
    X0, X1 = g_rect.x0, g_rect.x1
    maxw = X1 - X0 - 1.5  # 留 1.5pt 安全余量，防止 font.text_length 测量误差导致溢出
    avail = g_rect.y1 - g_rect.y0
    if from_text:
        en_lines = wrap(en_text, font, fs, maxw) if en_text.strip() else []
        cn_lines = wrap(cn_text, font, fs, maxw) if cn_text.strip() else []
        total = len(en_lines) + len(cn_lines)
        use_lh = lh if total * lh <= avail + 2 else avail / total
    else:
        # 空白模板：内容格很高，自动尽量放大字号铺满空间，避免底部大片留白
        en_lines = cn_lines = []
        for try_fs in (7.0, 6.2, 5.4, 4.6, 4.2):
            en_t = wrap(en_text, font, try_fs, maxw) if en_text.strip() else []
            cn_t = wrap(cn_text, font, try_fs, maxw) if cn_text.strip() else []
            en_lines, cn_lines, fs = en_t, cn_t, try_fs
            # 首行占约一个字号高，其余行各占 use_lh（最低 1.25 倍字号）
            n = len(en_lines) + len(cn_lines)
            if n and try_fs + (n - 1) * try_fs * 1.25 <= avail:
                break
        total = len(en_lines) + len(cn_lines)
        # 行距均匀铺满整格：总高 = 首行基线偏移 0.86fs + (n-1)*lh + 尾行下沉 0.29fs
        # （上限 1.9 倍字号，避免行与行之间过稀）
        use_lh = min((avail - fs * 1.15) / (total - 1), fs * 1.9) if total > 1 else lh
    y = g_rect.y0 + fs * 0.86
    tw = fitz.TextWriter(page.rect)
    for ln in en_lines + cn_lines:
        tw.append((X0, y), ln, font=font, fontsize=fs)
        y += use_lh
    tw.write_text(page, color=(0, 0, 0))
    print('goods lines:', len(en_lines), '+', len(cn_lines), 'bottom:', round(y, 1), '/', round(g_rect.y1, 1))

    if not a.no_ai_label:
        page.insert_text((W - 90, page.rect.height - 8), '内容由 AI 生成',
                         fontname="china-s", fontsize=8, color=(0.4, 0.4, 0.4))

    # 子集化嵌入字体（TextWriter 写中文会全量嵌入 SimSun，体积暴涨 10MB+）
    try:
        doc.subset_fonts()
    except Exception:
        pass
    doc.save(a.out, garbage=4, deflate=True)
    doc.close()

    # ---- 文本层校验 ----
    doc2 = fitz.open(a.out)
    txt = doc2[0].get_text()
    compact = re.sub(r'\s', '', txt)
    errors = []
    if old_no != a.new and old_no in txt:
        errors.append('旧运单号残留')
    if txt.count(a.new) != len(num_rects):
        errors.append(f'新运单号出现 {txt.count(a.new)} 次，期望 {len(num_rects)}')
    for cn, en in items:
        if re.sub(r'\s', '', cn.rstrip(',')) not in compact:
            errors.append('中文品名缺失: ' + cn[:20])
        if re.sub(r'\s', '', en.rstrip(',')) not in compact:
            errors.append('英文品名缺失: ' + en[:20])
    for tail in (a.en_tail, a.cn_tail):
        if tail and re.sub(r'\s', '', tail) not in compact:
            errors.append('结尾文本缺失: ' + tail[:20])
    if len(doc2) != len(fitz.open(a.template)):
        errors.append('页数不一致')
    doc2.close()
    if errors:
        print('VERIFY FAILED:')
        for e in errors:
            print(' -', e)
        sys.exit(1)
    print('VERIFY OK')


if __name__ == '__main__':
    main()
