#!/bin/bash
# DSV 服务器定期清理（配合 dsv-cleanup.timer，每 7 天一次）
# 清理范围：出货手续包/ 下的生成数据 + templates/ 下的测试*上传模板
# 关键约束：必须保留 _daemon_state.json（邮件去重记录），删除会导致重启后重复回信轰炸
set -u
PKG='/opt/dsv/出货手续包'
TPL='/opt/dsv/templates'
LOG=/var/log/dsv-cleanup.log

echo "===== $(date '+%F %T') 开始清理 =====" >> "$LOG"

systemctl stop dsv-mail
# 保留 _daemon_state.json（seen UID），其余全删（生成包/_mails/_存档/核对单/异常报告/任务文件）
find "$PKG" -mindepth 1 -maxdepth 1 ! -name '_daemon_state.json' -exec rm -rf {} + 2>>"$LOG"
mkdir -p "$PKG/_mails"
chown -R dsv:dsv "$PKG"
# 删除上传的测试模板（公共模板与 xlsx 不带"测试"前缀，不受影响）
# 排除长期保留的模板：测试176-61334210模板
find "$TPL" -mindepth 1 -maxdepth 1 -name '测试*' ! -name '测试176-61334210模板' -exec rm -rf {} + 2>>"$LOG"

systemctl start dsv-mail
sleep 2
if systemctl is-active --quiet dsv-mail; then
    echo "$(date '+%F %T') 清理完成，dsv-mail active" >> "$LOG"
else
    echo "$(date '+%F %T') 警告：清理后 dsv-mail 未 active！" >> "$LOG"
    exit 1
fi
