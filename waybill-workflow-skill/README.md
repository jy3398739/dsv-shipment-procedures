# 电池货出货手续包生成器（waybill-workflow）

从**订舱邮件**出发，一键生成一票电池货的完整出货手续包（7 份）：运单副本、安检单、应急措施、托运人声明、托书、交运单、危险品确认单。

## 技能包内容

```
waybill-workflow-skill/
├── SKILL.md                          # 技能定义（工作流说明）
├── README.md                         # 本说明
├── scripts/
│   ├── run_package.py                # ★ 一键入口：校验闸门 G1-G6 + 核对单/异常报告(HTML)
│   ├── mail_daemon.py                # ★ 服务器邮件守护：IMAP 收转发邮件→自动生成→回信 zip
│   ├── mail_config.ini.template      # 邮件守护配置模板（邮箱/授权码/公网地址）
│   ├── gen_procedures.py             # ★ 七份手续全套生成（主脚本，含 RUH 两主单数据）
│   ├── make_waybill.py               # 运单副本核心引擎（原位编辑 PDF）
│   ├── booking_to_waybill.py         # 订舱邮件 → 主单/分单/机型 → 运单副本
│   ├── make_procedures.py            # 订舱邮件 → 交运单 + 打包（旧版参考）
│   └── chinasdg_query.py             # 鉴定编号官网查询比对
├── server/
│   ├── dsv-mail.service              # systemd 服务单元（Linux 常驻）
│   └── 部署说明.md                    # ★ Linux 服务器部署完整步骤
└── references/
    └── 字段坐标与踩坑.md             # 七份手续字段坐标明细 + 实战踩坑

打包附带（`templates/`，本技能配套的模板与样例数据，解压后按需使用）：
- `templates/公共模板/`：空白基础模板（安检单.pdf 图片底 / 应急措施.pdf / 托运人声明.docx）
- `templates/测试2/`：RUH 样例票（176-61333915/176-61333926 的托书 + 运单副本）+ 订舱样例邮件 .eml
- `templates/测试176-61334210模板/`：176-61334210 填妥成品（字段坐标/版式参考）+ 交运单.xlsx 模板
- `templates/苹果手机-网站查询信息汇总.xlsx`：chinasdg 鉴定数据源（样例 55 条）
```

## 运行环境

- Python 3.12+（本机 `python` 已在 PATH；依赖 `PyMuPDF`、`openpyxl`，均已装）
  ```
  python -m pip install PyMuPDF openpyxl
  ```
- 安检单宋体写入：Windows 用系统 `simsun.ttc`；Linux 自动找 `templates/fonts/simsun.ttc` 或思源宋体/文泉驿（`song_font()` 按序探测）。
- 脚本已适配 PyMuPDF 1.28（insert_text 返回值/ToUnicode 坑，见 references/字段坐标与踩坑.md）。
- 路径均相对脚本位置定位，整个 DSV 目录可直接拷到 Linux 服务器运行。

## 需要的文件（数据源与模板）

| 用途 | 文件 | 说明 |
|---|---|---|
| 数据源 | 订舱邮件 `*.eml` | 提取主单/分单/机型/件重/航班 |
| 数据源 | `苹果手机-网站查询信息汇总.xlsx` | chinasdg 鉴定，sheet「网站查询信息」 |
| 空白模板 | `公共模板/` | 安检单.pdf（图片底）/ 应急措施.pdf / 2026国际出港电池货物托运人声明.docx |
| 票模板 | `测试2/{主单}模板/` | 该主单的托书.pdf + 运单副本.pdf（品名栏空） |
| 表单参考 | `测试176-61334210模板/` | 176-61334210 填妥成品（字段坐标/版式参考）+ 交运单.xlsx 模板 |

## 快速开始

```bash
# 0. 一键（脱离 AI 独立运行）：把订舱邮件 .eml 拖到 DSV 根目录的「一键出货手续包.bat」上
#    G2-G5 缺失 → 自动打开补充页（浏览器），在页面上直接补齐目的站/航班/日期、
#    分单件数毛重、机型证书编号/中英文品名，点「补充并重新校验」；
#    闸门全过 → 打开「生成前核对单」等人工确认，鉴定证书编号可在页面上直接改，
#    点「确认生成」才生成 出货手续包/{主单}/ 七份 + 正式核对单.html；点「取消」不生成；
#    邮件格式解析不了（G1）→ 出 异常报告.html 写明人工介入点。
#    （设 RUNPKG_AUTO_CONFIRM=1 可跳过确认页直接生成，自动化/测试用）

# 1. 从订舱邮件提取要素并匹配鉴定（查看结果）
python scripts/booking_to_waybill.py "订舱邮件.eml" --xlsx "苹果手机-网站查询信息汇总.xlsx"

# 1b. 邮件进 → 手续包出（提取聚 MASTERS，与内置核对后每主单生成七份）
python scripts/booking_to_waybill.py "订舱邮件.eml" --xlsx "苹果手机-网站查询信息汇总.xlsx" --full-package

# 2. 生成七份手续（gen_procedures.py 内置主单数据，改 MASTERS 即可换票）
python scripts/gen_procedures.py
# 输出：出货手续包/{主单}/ 下 7 份文件

# 3. 单张运单副本（自由指定）
python scripts/make_waybill.py --template "旧运单.pdf" --old 176-61334210 --new 176-99999999 \
    --excel "苹果手机-网站查询信息汇总.xlsx" --out "新运单.pdf" --no-ai-label
```

## 换新票时改什么

1. 在 `gen_procedures.py` 的 `MASTERS` 字典填入新主单：机型列表、电池型号、件数、分单(HAWB,件数,重量)。
2. 确保 `测试2/{主单}模板/` 有该主单的**托书 + 运单副本**。
3. 若机型不在 xlsx，先补查 chinasdg 并写入（或补进 `CN_GOODS`/`EN_GOODS`）。
4. 运行 `gen_procedures.py`。

## 服务器部署（同事转发邮件 → 自动生成 → 回信）

完整步骤见 `server/部署说明.md`。要点：

1. 整个 DSV 目录传到服务器（如 `/opt/dsv/`），装 `pymupdf python-docx openpyxl` + `fonts-noto-cjk`。
2. `scripts/mail_config.ini`（照 template 填）：IMAP/SMTP 邮箱 + 授权码 + `public_base`（公网 IP:端口）。
3. systemd 装 `server/dsv-mail.service`，放行 `http_port` 端口。
4. 同事转发订舱邮件 → 60 秒内处理：闸门全过直接回信 zip+核对单；缺信息回信补充页链接（公网令牌页，补齐提交后自动生成回信）；解析不了回信异常说明。
5. 同主单再次转发 = 按最新邮件重新生成，旧版自动存档到 `出货手续包/_存档/`；新主单的托书/运单副本模板也是发邮件：作为附件发到收件邮箱（文件名含主单号+托书/运单副本）自动入库，可随订舱邮件一起发或单独发。

## 关键约定

- 品名变体统一用「数字式手机 / 5G数字移动电话机」+「DIGITAL CELLPHONE / 5G DIGITAL MOBILE PHONE FOR RADIOTELEPHONY」（与参考版 176-61334210 一致）。
- 安检单品名 5.2pt 窄格、应急措施 7pt、托书英文 5pt + 中文 4.8pt（**完全照参考版**，不要自行放大字号或扩格）。
- 运单副本只填品名区，不动运单号/条码/件重。
- 日期格式 `2026-08-16`；航班如 EK309；目的站 RUH。
