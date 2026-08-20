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
    """右半页小字号连续行段，返回 (rect, fontname, fontsize, line_h)。"""
    spans = []
    for b in page.get_text('dict')['blocks']:
        for l in b.get('lines', []):
            for s in l['spans']:
                if s['size'] <= 8 and s['bbox'][0] > page.rect.width * 0.6:
                    spans.append(s)
    spans.sort(key=lambda s: s['bbox'][1])
    lines = []
    for s in spans:
        if lines and abs(s['bbox'][1] - lines[-1][-1]['bbox'][1]) < 2:
            lines[-1].append(s)
        else:
            lines.append([s])
    # 最长连续段（行距 < 8）
    segs, cur = [], [lines[0]]
    for ln in lines[1:]:
        if ln[0]['bbox'][1] - cur[-1][0]['bbox'][1] < 8:
            cur.append(ln)
        else:
            segs.append(cur)
            cur = [ln]
    segs.append(cur)
    seg = max(segs, key=len)
    flat = [s for ln in seg for s in ln]
    x0 = min(s['bbox'][0] for s in flat)
    x1 = max(s['bbox'][2] for s in flat)
    y0 = min(s['bbox'][1] for s in flat)
    y1 = max(s['bbox'][3] for s in flat)
    fs = statistics.median(s['size'] for s in flat)
    lhs = [seg[i + 1][0]['bbox'][1] - seg[i][0]['bbox'][1] for i in range(len(seg) - 1)]
    lh = statistics.median(lhs) if lhs else fs * 1.2
    # 上下界不得越过表格横线
    hlines = []
    for d in page.get_drawings():
        for it in d['items']:
            if it[0] == 'l':
                p0, p1 = it[1], it[2]
                if abs(p0.y - p1.y) < 0.5 and (p1.x - p0.x) > page.rect.width * 0.5 and p0.x < x0:
                    hlines.append(p0.y)
    top = max([h for h in hlines if h < y0], default=y0 - 2)
    bot = min([h for h in hlines if h > y1], default=y1 + 2)
    rect = fitz.Rect(x0 - 4, max(top + 1, y0 - 2), min(x1 + 4, page.rect.width - 10), min(bot - 1, y1 + 2))
    main = max(flat, key=lambda s: len(s['text']))
    return rect, main['font'], fs, lh


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
    units = re.findall(r'[A-Za-z0-9][A-Za-z0-9.,:()+/\-]*|\s|.', text)
    lines, cur = [], ''

    def force_split(s):
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
            if cur.strip():
                lines.append(cur.rstrip())
            cur = '' if u == ' ' else u
            if font.text_length(cur, fontsize=fs) > maxw:
                extra, cur = force_split(cur)
                lines.extend(extra)
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
    en_text = a.en_head + ",".join(en for _, en in items) + (("." + a.en_tail) if a.en_tail else "")
    cn_text = ",".join(cn for cn, _ in items) + a.cn_tail

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
    g_rect, g_fontname, fs, lh = find_goods_area(page)
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
    X0, X1 = g_rect.x0, min(g_rect.x1, W - 12)
    en_lines = wrap(en_text, font, fs, X1 - X0) if en_text.strip() else []
    cn_lines = wrap(cn_text, font, fs, X1 - X0) if cn_text.strip() else []
    total = len(en_lines) + len(cn_lines)
    avail = g_rect.y1 - g_rect.y0
    use_lh = lh if total * lh <= avail + 2 else avail / total
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
        if re.sub(r'\s', '', cn) not in compact:
            errors.append('中文品名缺失: ' + cn[:20])
        if re.sub(r'\s', '', en) not in compact:
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
