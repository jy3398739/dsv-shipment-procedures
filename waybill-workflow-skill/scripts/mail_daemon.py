# -*- coding: utf-8 -*-
"""邮件守护：服务器收到同事转发的订舱邮件 → 自动跑手续包管线 → 回信结果。

流程：
  IMAP 轮询收件箱 UNSEEN → 落盘 .eml（转发件自动剥出内嵌原始邮件）→
  附件里的「主单号+托书/运单副本 .pdf」自动放进 templates/测试2/{主单}模板/ →
  run_package.run_headless 探测闸门：
    hard      → 回信异常报告（需人工处理）
    need_fill → 回信带令牌补充页链接（公网地址），同事浏览器补齐提交后自动生成并回信
    ready     → 直接生成：同主单旧手续包存档到 出货手续包/_存档/ → 打 zip → 回信
同一主单再次转发 = 覆盖重生成（旧版进存档），应对「一封邮件不准、后期更改」。

用法：python mail_daemon.py [配置文件路径]（默认同目录 mail_config.ini）
依赖：仅标准库（imaplib/smtplib/email/http.server）+ 管线依赖。
"""
import configparser
import datetime
import email
import email.utils
import http.server
import imaplib
import json
import os
import re
import secrets
import shutil
import smtplib
import sys
import threading
import zipfile
from email.header import decode_header
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from urllib.parse import urlparse

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import run_package as rp
import booking_to_waybill as bw
import gen_procedures as gp

OUT = rp.OUT
WORK = os.path.join(OUT, '_mails')          # 收到的邮件落盘
ARCHIVE = os.path.join(OUT, '_存档')        # 重生成前的旧版手续包
STATE_F = os.path.join(OUT, '_daemon_state.json')  # 已处理 UID，防重启重复处理

MASTER_RE = re.compile(r'176-\d{8}')


def log(*a):
    print(datetime.datetime.now().strftime('%H:%M:%S'), *a, flush=True)


# ================== 配置 ==================
def load_config(path):
    if not os.path.isfile(path):
        print(f"找不到配置文件 {path}，参照 mail_config.ini.template 建一个。")
        sys.exit(2)
    cp = configparser.ConfigParser()
    cp.read(path, encoding='utf-8')
    c = {
        'imap_host': cp.get('mail', 'imap_host'),
        'imap_port': cp.getint('mail', 'imap_port', fallback=993),
        'smtp_host': cp.get('mail', 'smtp_host'),
        'smtp_port': cp.getint('mail', 'smtp_port', fallback=465),
        'user': cp.get('mail', 'user'),
        'password': cp.get('mail', 'password'),  # QQ/163 用授权码
        'poll_sec': cp.getint('mail', 'poll_sec', fallback=60),
        'filter_keywords': cp.get('mail', 'filter_keywords', fallback=''),
        'filter_froms': cp.get('mail', 'filter_froms', fallback=''),
        'send_to': cp.get('server', 'send_to', fallback=''),      # 空=回给发件人
        'public_base': cp.get('server', 'public_base').rstrip('/'),
        'http_port': cp.getint('server', 'http_port', fallback=8080),
    }
    return c


# ================== 状态（已处理 UID） ==================
def load_state():
    try:
        return json.load(open(STATE_F, encoding='utf-8'))
    except Exception:
        return {'seen': []}


def save_state(state):
    os.makedirs(OUT, exist_ok=True)
    json.dump(state, open(STATE_F, 'w', encoding='utf-8'), ensure_ascii=False)
    state['seen'] = state['seen'][-2000:]


# ================== 邮件收取与拆解 ==================
def hdr(msg, name):
    v = msg.get(name) or ''
    try:
        parts = decode_header(v)
        out = []
        for s, cs in parts:
            out.append(s.decode(cs or 'utf-8', 'ignore') if isinstance(s, bytes) else s)
        return ''.join(out)
    except Exception:
        return v


def _dec_header(s):
    """RFC2047 编码的文件名（=?utf-8?B?...?=）解码成明文。"""
    if isinstance(s, bytes):
        s = s.decode('utf-8', 'ignore')
    if not s or '=?' not in s:
        return s or ''
    from email.header import decode_header
    try:
        return ''.join(p.decode(ch or 'utf-8', 'ignore') if isinstance(p, bytes) else str(p)
                       for p, ch in decode_header(s))
    except Exception:
        return s


def save_template_file(fname, payload):
    """按文件名规则（主单号+托书/运单副本）把模板 PDF 入库。成功返回入库文件名，否则 None。"""
    fname = _dec_header(fname) or ''
    m = MASTER_RE.search(fname)
    if not m or not re.search(r'托书|运单副本|运单复本', fname):
        return None
    master = m.group(0)
    tdir = os.path.join(gp.TMPL2, master + '模板')
    os.makedirs(tdir, exist_ok=True)
    if not payload:
        return None
    # 同名覆盖（同事重发更新模板）
    clean = re.sub(r'[/\\:*?"<>|]', '_', fname)
    open(os.path.join(tdir, clean), 'wb').write(payload)
    log(f"  [模板] {clean} → {os.path.basename(tdir)}/")
    return clean


def save_attachment_templates(msg):
    """附件名含 主单号+托书/运单副本 的 pdf → templates/测试2/{主单}模板/。返回入库文件名列表。"""
    saved = []
    for part in msg.walk():
        fname = _dec_header(part.get_filename()) or ''
        if part.get_content_type() != 'application/pdf' and \
                not fname.lower().endswith('.pdf'):
            continue
        r = save_template_file(fname, part.get_payload(decode=True))
        if r:
            saved.append(r)
    return saved


def extract_eml_files(msg, stamp):
    """从收到的邮件里取出可解析的订舱邮件本体，返回 .eml 路径列表。
    转发邮件两种形态都覆盖：内嵌 message/rfc822 附件 / 邮件本身就是正文带表格。"""
    os.makedirs(WORK, exist_ok=True)
    found = []
    for part in msg.walk():
        if part.get_content_type() == 'message/rfc822':
            subs = part.get_payload()
            if not isinstance(subs, list):
                subs = [subs]
            for i, sub in enumerate(subs):
                p = os.path.join(WORK, f'{stamp}_内嵌{i}.eml')
                open(p, 'wb').write(sub.as_bytes())
                found.append(p)
        else:
            fname = part.get_filename() or ''
            payload = part.get_payload(decode=True)
            if not payload:
                continue
            if fname.lower().endswith('.eml'):
                p = os.path.join(WORK, f'{stamp}_{re.sub(r"[/\\:*?\"<>|]", "_", fname)}')
                open(p, 'wb').write(payload)
                found.append(p)
            elif part.get_content_type() == 'application/octet-stream' and \
                    payload.lstrip()[:20].split(b':', 1)[0].strip().lower() in \
                    (b'from', b'received', b'return-path', b'date', b'subject', b'message-id'):
                # 部分客户端把 .eml 附件标成 octet-stream 且丢文件名，按邮件头嗅探
                p = os.path.join(WORK, f'{stamp}_附件嗅探{len(found)}.eml')
                open(p, 'wb').write(payload)
                found.append(p)
    # 没有内嵌件：把整封邮件自己当订舱邮件（直接转发的正文形态）
    self_p = os.path.join(WORK, f'{stamp}_整封.eml')
    open(self_p, 'wb').write(msg.as_bytes())
    found.append(self_p)
    return found


def pick_parseable(emls):
    """返回正文能解析出表格块的 .eml；都解析不了返回空列表（调用方决定怎么处理）。"""
    for p in emls:
        try:
            if bw.parse_email(p):
                return [p]
        except Exception:
            continue
    return []


# ================== 回信 ==================
def smtp_send(cfg, to, subject, html, files=(), in_reply_to=None):
    msg = MIMEMultipart()
    msg['From'] = cfg['user']
    msg['To'] = to
    msg['Subject'] = subject
    if in_reply_to:
        msg['In-Reply-To'] = in_reply_to
        msg['References'] = in_reply_to
    msg.attach(MIMEText(html, 'html', 'utf-8'))
    for path in files:
        with open(path, 'rb') as f:
            att = MIMEApplication(f.read())
        att.add_header('Content-Disposition', 'attachment',
                       filename=('utf-8', '', os.path.basename(path)))
        msg.attach(att)
    if cfg['smtp_port'] == 465:
        s = smtplib.SMTP_SSL(cfg['smtp_host'], cfg['smtp_port'], timeout=60)
    else:
        s = smtplib.SMTP(cfg['smtp_host'], cfg['smtp_port'], timeout=60)
        s.starttls()
    s.login(cfg['user'], cfg['password'])
    s.sendmail(cfg['user'], [to], msg.as_string())
    s.quit()
    log(f"  [回信] → {to}：{subject}")


def reply_addr(cfg, msg):
    to = cfg['send_to'].strip()
    if to:
        return to
    r = hdr(msg, 'Reply-To').strip() or hdr(msg, 'From')
    return email.utils.parseaddr(r)[1]


# ================== 存档与打包 ==================
def archive_existing(masters):
    """同主单重复转发：旧手续包挪进 _存档/{主单}_{时间}/，保证拿到的是最新一版。"""
    for master in masters:
        d = os.path.join(OUT, master)
        if os.path.isdir(d):
            ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            dst = os.path.join(ARCHIVE, f'{master}_{ts}')
            os.makedirs(ARCHIVE, exist_ok=True)
            shutil.move(d, dst)
            log(f"  [存档] 旧版 {master} → {os.path.basename(dst)}/")


def zip_masters(masters):
    """把各主单手续包目录打一个 zip，返回路径。"""
    stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    zp = os.path.join(WORK, f'手续包_{stamp}.zip')
    os.makedirs(WORK, exist_ok=True)
    with zipfile.ZipFile(zp, 'w', zipfile.ZIP_DEFLATED) as z:
        for master in masters:
            d = os.path.join(OUT, master)
            for f in sorted(os.listdir(d)):
                z.write(os.path.join(d, f), f'{master}/{f}')
    return zp


def zip_single_master(master):
    """把单个主单的手续包目录打一个 zip，返回路径。"""
    stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    zp = os.path.join(WORK, f'{master}_{stamp}.zip')
    os.makedirs(WORK, exist_ok=True)
    d = os.path.join(OUT, master)
    with zipfile.ZipFile(zp, 'w', zipfile.ZIP_DEFLATED) as z:
        for f in sorted(os.listdir(d)):
            z.write(os.path.join(d, f), f)
    return zp


# ================== 结果邮件 HTML ==================
def result_html(ok, masters, report_html):
    color = '#1a7f37' if ok else '#b42318'
    tip = ('手续包已生成（见附件 zip，解压即用）。逐项校验结果如下，'
           if ok else '手续包已生成但<b>校验有不通过项</b>，请勿直接使用，先人工复核：')
    return f"""<html><body style="font-family:'Microsoft YaHei',sans-serif">
<div style="padding:10px 14px;color:#fff;background:{color};border-radius:6px;font-size:15px">
出货手续包 {'全部校验通过' if ok else '存在异常'}　|　主单：{', '.join(masters)}（{len(masters)} 个 × 7 份）</div>
<p>{tip}</p>
{report_html}
<p style="color:#666;font-size:12px">本邮件由出货手续包服务器自动生成。同主单再次转发邮件 = 按最新内容重新生成（旧版自动存档）。</p>
</body></html>"""


def confirm_mail_html(base_url, token, masters):
    url = f'{base_url}/t/{token}/confirm'
    return f"""<html><body style="font-family:'Microsoft YaHei',sans-serif">
<div style="padding:10px 14px;color:#fff;background:#1a7f37;border-radius:6px;font-size:15px">
订舱邮件信息齐全，待人工确认生成手续包　|　主单：{', '.join(masters)}（{len(masters)} 个 × 7 份）</div>
<p>请打开下面的链接，核对各主单 HAWB/件重/机型/鉴定证书编号（可直接修改）后点「确认生成」，手续包才会生成并回信：</p>
<p><a href="{url}" style="font-size:16px">{url}</a></p>
<p style="color:#666;font-size:12px">未确认不会生成任何文件；点「取消」则不生成。链接长期有效。</p>
</body></html>"""


def fill_mail_html(base_url, token, issues):
    url = f'{base_url}/t/{token}/'
    li = ''.join(f'<li>{rp.esc(i)}</li>' for i in issues)
    return f"""<html><body style="font-family:'Microsoft YaHei',sans-serif">
<div style="padding:10px 14px;color:#fff;background:#b45309;border-radius:6px;font-size:15px">
订舱邮件有 {len(issues)} 项信息缺失，需要补充后才能生成手续包</div>
<p>请打开下面的链接，在网页里补齐红色框项后点「补充并生成」，手续包会自动生成并回信：</p>
<p><a href="{url}" style="font-size:16px">{url}</a></p>
<ul>{li}</ul>
<p style="color:#666;font-size:12px">链接长期有效，随时可打开继续填写。</p>
</body></html>"""


def hard_mail_html(issues):
    li = ''.join(f'<li>{rp.esc(i)}</li>' for i in issues)
    return f"""<html><body style="font-family:'Microsoft YaHei',sans-serif">
<div style="padding:10px 14px;color:#fff;background:#b42318;border-radius:6px;font-size:15px">
订舱邮件无法自动处理，需人工介入</div>
<ul>{li}</ul>
<p style="color:#666;font-size:12px">常见原因：转发时正文变成图片/附件、邮件格式变化。
可把原始订舱邮件作为附件转发，或直接转发纯文本正文。</p>
</body></html>"""


def templates_received_html(saved):
    li = ''.join(f'<li>{rp.esc(f)}</li>' for f in saved)
    return f"""<html><body style="font-family:'Microsoft YaHei',sans-serif">
<div style="padding:10px 14px;color:#fff;background:#1a7f37;border-radius:6px;font-size:15px">
已收到 {len(saved)} 份模板文件，并自动放入对应主单模板库</div>
<ul>{li}</ul>
<p>若有正在等待该模板的手续包任务，稍后会自动生成并把结果回信给您；无需其他操作。</p>
<p style="color:#666;font-size:12px">模板要求：PDF 附件，文件名包含主单号（如 176-61334567）+「托书」或「运单副本」；
同名重发会覆盖旧模板。</p>
</body></html>"""


def server_issues(issues):
    """服务器模式改写问题文案：G5 缺模板从「手动放目录」改成「邮件发附件」。"""
    out = []
    for i in issues:
        if '缺模板' in i:
            m = MASTER_RE.search(i)
            mst = m.group(0) if m else '该主单'
            out.append(f"G5 主单 {mst} 缺托书/运单副本模板：把这两份 PDF 作为邮件附件发到本邮箱"
                       f"（文件名含主单号，如「{mst}托书.pdf」）即可自动入库；"
                       f"入库后本页点「补充并重新校验」，或直接等自动生成回信。")
        else:
            out.append(i)
    return out


# ================== 令牌补充页（HTTP，任务落盘防重启丢链接） ==================
JOBS = {}  # token -> {'st':..., 'issues':..., 'to':..., 'msg_id':..., 'subject':...}
JOBS_LOCK = threading.Lock()
JOBS_F = os.path.join(OUT, '_daemon_jobs.json')
CFG = None


def _sanitize(obj):
    """递归清理：Message 等不可序列化对象转成可落盘占位（生成阶段用不到 msg）。"""
    import email as _email
    if isinstance(obj, _email.message.Message):
        return {'_msg_placeholder': obj.get('Subject', '')[:80]}
    if isinstance(obj, dict):
        return {k: _sanitize(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_sanitize(v) for v in obj]
    if isinstance(obj, (set, frozenset)):
        return [_sanitize(v) for v in obj]
    return obj


def jobs_dump():
    os.makedirs(OUT, exist_ok=True)
    with JOBS_LOCK:
        snap = dict(JOBS)
    try:
        tmp = JOBS_F + '.tmp'
        json.dump(_sanitize(snap), open(tmp, 'w', encoding='utf-8'), ensure_ascii=False)
        os.replace(tmp, JOBS_F)
    except Exception as e:
        log(f"  [警告] 补充任务存档失败：{e}")


def jobs_load():
    try:
        data = json.load(open(JOBS_F, encoding='utf-8'))
        with JOBS_LOCK:
            JOBS.update(data)
        if data:
            log(f"[恢复] 载入 {len(data)} 个待补充任务（旧补充链接继续有效）")
    except Exception:
        pass


class TokenHandler(http.server.BaseHTTPRequestHandler):
    protocol_version = 'HTTP/1.1'
    timeout = 60

    def log_message(self, fmt, *args):
        pass

    def _send(self, code, body, ctype='text/html; charset=utf-8', binary=False):
        data = body if binary else body.encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(data)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(data)

    def _path(self):
        """nginx 反代前缀（如 /dsv），用于页面 JS 拼接 fetch 绝对路径。"""
        return urlparse(CFG['public_base']).path.rstrip('/')

    def do_GET(self):
        m = re.match(r'^/t/([\w-]+)(/confirm)?/?$', self.path)
        if not m:
            return self._send(404, 'not found', 'text/plain')
        with JOBS_LOCK:
            job = JOBS.get(m.group(1))
        if not job:
            return self._send(404, '该链接不存在或已处理完成', 'text/plain')
        base = self._path() + f'/t/{m.group(1)}'
        if m.group(2):  # /confirm 生成前核对单（人工确认）
            page = rp.build_confirm_html(job['st'], 0, base_url=base)
            return self._send(200, page)
        page = rp.build_form_html(
            job['st'], job['issues'], 0,
            ok_msg='校验通过，即将打开「生成前核对单」——请核对信息后点「确认生成」，完成后结果将发到您的邮箱。',
            base_url=base)
        self._send(200, page)

    def do_POST(self):
        m = re.match(r'^/t/([\w-]+)/(submit|confirm_submit|cancel|upload|back)$', self.path)
        if not m:
            return self._send(404, 'not found', 'text/plain')
        token, action = m.group(1), m.group(2)
        if action == 'back':
            # 从核对单返回补充页，不需要额外处理，JS 会自动跳转
            return self._send(200, json.dumps({'back': True}, ensure_ascii=False), 'application/json')
        with JOBS_LOCK:
            job = JOBS.get(token)
        if not job:
            return self._send(404, json.dumps({'ok': False, 'issues': ['链接不存在或已处理完成']},
                                              ensure_ascii=False), 'application/json')
        st = job['st']
        n = int(self.headers.get('Content-Length') or 0)

        if action == 'cancel':
            with JOBS_LOCK:
                JOBS.pop(token, None)
            jobs_dump()
            return self._send(200, json.dumps({'cancelled': True}, ensure_ascii=False), 'application/json')

        if action == 'upload':
            # 补充页上传模板 PDF（multipart/form-data）→ 入库 → 重跑闸门
            import email as _email
            from email import policy as _policy
            ct = self.headers.get('Content-Type') or ''
            saved = []
            try:
                fake = ('Content-Type: ' + ct + '\r\nMIME-Version: 1.0\r\n\r\n').encode() + self.rfile.read(n)
                up = _email.message_from_bytes(fake, policy=_policy.default)
                for part in up.iter_parts():
                    fn = part.get_filename()
                    if not fn:
                        continue
                    payload = part.get_payload(decode=True)
                    if payload and _dec_header(fn).lower().endswith('.pdf'):
                        r = save_template_file(fn, payload)
                        if r:
                            saved.append(r)
            except Exception as e:
                return self._send(400, json.dumps({'ok': False, 'issues': [f'上传解析失败：{e}']},
                                                  ensure_ascii=False), 'application/json')
            try:
                st['certs'] = bw.load_certs(rp.XLSX)
            except Exception:
                pass
            rp.rebuild_state(st)
            remain = server_issues(rp.collect_gate_issues(st))
            job['issues'] = remain
            jobs_dump()
            resp = {'ok': True, 'saved': saved}
            if remain:
                resp['remain'] = remain
            else:
                resp['confirm_url'] = self._path() + f'/t/{token}/confirm'
            return self._send(200, json.dumps(resp, ensure_ascii=False), 'application/json')

        # submit / confirm_submit 走 JSON body（cancel/upload 已提前 return）
        try:
            patch = json.loads(self.rfile.read(n).decode('utf-8')) if n else {}
        except Exception:
            return self._send(400, json.dumps({'ok': False, 'issues': ['补充输入解析失败']},
                                              ensure_ascii=False), 'application/json')
        try:
            st['certs'] = bw.load_certs(rp.XLSX)
        except Exception:
            pass
        if action == 'confirm_submit':
            # 确认页提交：证书编号可改，确认后才生成+回信
            notes = rp.apply_cert_edits(st, patch.get('certs') or {})
            rp.rebuild_state(st)
            remain = server_issues(rp.collect_gate_issues(st))
            if remain:
                job['issues'] = remain
                jobs_dump()
                return self._send(200, json.dumps({'ok': False, 'issues': remain},
                                                  ensure_ascii=False), 'application/json')
            err = finish_job(token)
            if err:
                return self._send(200, json.dumps({'ok': False, 'issues': [err]},
                                                  ensure_ascii=False), 'application/json')
            self._send(200, json.dumps({'ok': True, 'notes': notes}, ensure_ascii=False),
                       'application/json')
            return
        rp.apply_patch(st, patch)
        rp.rebuild_state(st)
        remain = server_issues(rp.collect_gate_issues(st))
        if remain:
            job['issues'] = remain
            jobs_dump()
            return self._send(200, json.dumps({'ok': False, 'issues': remain},
                                              ensure_ascii=False), 'application/json')
        # 闸门全过 → 打开生成前核对单（人工确认后才生成）
        self._send(200, json.dumps(
            {'ok': False, 'confirm': True,
             'confirm_url': self._path() + f'/t/{token}/confirm'},
            ensure_ascii=False), 'application/json')


def finish_job(token):
    """任务闸门全过后：存档旧版→生成→回信（每主单单独 zip + 核对单）→移除任务。"""
    with JOBS_LOCK:
        job = JOBS.get(token)
    if not job:
        return '任务不存在'
    try:
        st = job['st']
        masters = list(st['masters_all'].keys())
        archive_existing(masters)
        all_checks, bad, ok, rp_path = rp.finish(st, job['subject'] or '邮件补充')
        # 每个主单单独打一个 zip，加核对单，全部挂在一封邮件的附件里
        attachments = [rp_path]
        for master in masters:
            zp = zip_single_master(master)
            attachments.append(zp)
        smtp_send(CFG, job['to'],
                  f"出货手续包已生成（{len(masters)} 个主单 × 7 份）{'（有异常项请复核）' if not ok else ''}",
                  result_html(ok, masters,
                              f'<p>手续包已生成，每个主单单独一个 zip 附件，解压即用。核对单见附件。</p>'),
                  files=attachments, in_reply_to=job['msg_id'])
    except Exception as e:
        import traceback
        traceback.print_exc()
        return f'生成报错：{e}'
    with JOBS_LOCK:
        JOBS.pop(token, None)
    jobs_dump()
    return None


def retry_pending_jobs(cfg):
    """模板/信息可能有变（收到模板附件、xlsx 更新）时，把待补充任务重新过一遍闸门，
    能过的直接生成回信（同事不用再去点补充页）。"""
    with JOBS_LOCK:
        toks = list(JOBS)
    for tok in toks:
        job = JOBS.get(tok)
        if not job:
            continue
        st = job['st']
        try:
            st['certs'] = bw.load_certs(rp.XLSX)
        except Exception:
            pass
        rp.rebuild_state(st)
        remain = server_issues(rp.collect_gate_issues(st))
        if remain:
            job['issues'] = remain  # 补充页显示最新缺项
            continue
        log(f"  [自动补齐] 任务 /t/{tok}/ 闸门已过（模板/信息到位），直接生成回信")
        finish_job(tok)
    jobs_dump()


def serve_tokens(port):
    srv = http.server.ThreadingHTTPServer(('0.0.0.0', port), TokenHandler)
    srv.daemon_threads = True
    log(f"[HTTP] 令牌补充页服务已启动：0.0.0.0:{port}")
    srv.serve_forever()


# ================== 单封邮件处理 ==================
# 服务自己回信的 SMTP 副本主题前缀（QQ 自收自发会进收件箱，误判为订舱邮件会无限循环）
REPLY_PREFIXES = ('订舱邮件无法自动处理', '手续包需要补充', '出货手续包已生成',
                  '模板已收到并入库', '请确认生成手续包')


def should_process(cfg, msg):
    """过滤：主题含关键词或发件人命中白名单才处理；两者都未配置=全处理（原行为）。"""
    frm = hdr(msg, 'From') or ''
    subj = hdr(msg, 'Subject') or ''
    if any(subj.startswith(p) for p in REPLY_PREFIXES):
        return False
    wl = [w.strip().lower() for w in cfg.get('filter_froms', '').split(',') if w.strip()]
    if wl and any(w in frm.lower() for w in wl):
        return True
    kw = [k.strip().lower() for k in cfg.get('filter_keywords', '').split(',') if k.strip()]
    if kw and any(k in subj.lower() for k in kw):
        return True
    return not kw and not wl


def process_message(cfg, msg):
    """处理一封收到的邮件：落盘→模板附件→跑管线→按结果回信。"""
    subject = hdr(msg, 'Subject') or '(无主题)'
    frm = hdr(msg, 'From')
    msg_id = hdr(msg, 'Message-ID')
    stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S') + '_' + secrets.token_hex(3)
    to = reply_addr(cfg, msg)
    log(f"[邮件] {frm} | {subject[:40]}")

    saved = save_attachment_templates(msg)
    candidates = extract_eml_files(msg, stamp)
    emls = pick_parseable(candidates)
    if not emls:
        if saved:
            # 纯模板邮件（没有订舱内容）：先回确认，再尝试自动完成等模板的待补充任务
            log(f"  [结果] 纯模板邮件，{len(saved)} 份已入库")
            smtp_send(cfg, to, f"模板已收到并入库（{len(saved)} 份）",
                      templates_received_html(saved), in_reply_to=msg_id)
            retry_pending_jobs(cfg)
            return
        emls = [candidates[-1]]  # 无模板也无订舱内容：让管线出具体硬伤说明

    res = rp.run_headless(emls, generate=False)
    base = cfg['public_base']
    if res['status'] == 'hard':
        log("  [结果] 硬伤，回信异常说明")
        smtp_send(cfg, to, f"订舱邮件无法自动处理：{subject[:30]}",
                  hard_mail_html(res['issues']), in_reply_to=msg_id)
        return
    if res['status'] == 'need_fill':
        issues = server_issues(res['issues'])
        token = secrets.token_urlsafe(12)
        with JOBS_LOCK:
            JOBS[token] = {'st': res['st'], 'issues': issues, 'to': to,
                           'msg_id': msg_id, 'subject': subject}
        jobs_dump()
        log(f"  [结果] 缺 {len(issues)} 项，回信补充链接 /t/{token}/")
        smtp_send(cfg, to, f"手续包需要补充 {len(issues)} 项信息：{subject[:30]}",
                  fill_mail_html(base, token, issues), in_reply_to=msg_id)
        return
    # ready：信息齐全 → 回信确认链接，人工确认后才生成
    st = res['st']
    token = secrets.token_urlsafe(12)
    masters = list(st['masters_all'].keys())
    with JOBS_LOCK:
        JOBS[token] = {'st': st, 'issues': [], 'to': to,
                       'msg_id': msg_id, 'subject': subject}
    jobs_dump()
    log(f"  [结果] 信息齐全，回信确认链接 /t/{token}/confirm")
    smtp_send(cfg, to, f"请确认生成手续包（{len(masters)} 个主单）：{subject[:30]}",
              confirm_mail_html(base, token, masters), in_reply_to=msg_id)


# ================== IMAP 轮询 ==================
def poll_loop(cfg, state):
    while True:
        try:
            box = imaplib.IMAP4_SSL(cfg['imap_host'], cfg['imap_port'])
            box.login(cfg['user'], cfg['password'])
            log(f"[IMAP] 已连接 {cfg['imap_host']}，开始轮询（每 {cfg['poll_sec']}s）")
            box.select('INBOX')
            while True:
                # 搜最近 3 天全部邮件（含已读）：QQ 邮箱会把转发件标成已读，
                # 只看 UNSEEN 会漏掉 → 用 seen 列表去重，已读未处理的一样处理
                since = (datetime.date.today() - datetime.timedelta(days=3)).strftime('%d-%b-%Y')
                typ, data = box.uid('search', None, 'SINCE', since)
                uids = (data[0] or b'').split() if typ == 'OK' else []
                for uid in uids:
                    uid_s = uid.decode()
                    if uid_s in state['seen']:
                        box.uid('store', uid, '+FLAGS', '\\Seen')
                        continue
                    typ2, md = box.uid('fetch', uid, '(RFC822)')
                    if typ2 != 'OK' or not md or not md[0]:
                        continue
                    try:
                        msg = email.message_from_bytes(md[0][1])
                        if not should_process(cfg, msg):
                            log(f"  [跳过] 非订舱邮件（无关键词/白名单命中）：{(hdr(msg, 'Subject') or '')[:40]}")
                            state['seen'].append(uid_s)
                            save_state(state)
                            box.uid('store', uid, '+FLAGS', '\\Seen')
                            continue
                        process_message(cfg, msg)
                    except Exception as e:
                        import traceback
                        traceback.print_exc()
                        log(f"  [错误] 处理失败：{e}")
                    state['seen'].append(uid_s)
                    save_state(state)
                    box.uid('store', uid, '+FLAGS', '\\Seen')
                import time
                time.sleep(cfg['poll_sec'])
        except KeyboardInterrupt:
            raise
        except Exception as e:
            log(f"[IMAP] 连接断开/出错：{e}，{cfg['poll_sec']}s 后重连")
            import time
            time.sleep(cfg['poll_sec'])


def main():
    global CFG
    cfg_path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, 'mail_config.ini')
    CFG = load_config(cfg_path)
    base = urlparse(CFG['public_base'])
    if not base.scheme or not base.netloc:
        print(f"配置 public_base 不是完整网址（当前：{CFG['public_base']}）")
        sys.exit(2)
    state = load_state()
    jobs_load()
    threading.Thread(target=serve_tokens, args=(CFG['http_port'],), daemon=True).start()
    poll_loop(CFG, state)


if __name__ == '__main__':
    main()
