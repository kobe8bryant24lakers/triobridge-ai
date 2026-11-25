#!/bin/bash
set -euo pipefail
cd "$(dirname "$0")"
apps="$HOME/Applications"
services="$HOME/Library/Services"
mkdir -p "$apps" "$services"
app="$apps/TrioBridge AI.app"
provider="$services/TrioBridge AI.service"
verify_bundle() {
  [ ! -e "$1" ] || [ "$(/usr/libexec/PlistBuddy -c 'Print CFBundleIdentifier' "$1/Contents/Info.plist")" = "$2" ]
}
verify_bundle "$app" ai.triobridge.workbench
verify_bundle "$provider" ai.triobridge.workbench.service
verify_bundle 'TrioBridge AI.app' ai.triobridge.workbench
verify_bundle 'TrioBridge AI.service' ai.triobridge.workbench.service
if /usr/bin/pgrep -x 'TrioBridge AI' >/dev/null; then
  echo '请先关闭三桥 AI 工作台，再重新双击安装。'; exit 1
fi
stage=$(/usr/bin/mktemp -d "$apps/.triobridge-install.XXXXXX")
trap '/bin/rm -rf "$stage"' EXIT
/usr/bin/ditto 'TrioBridge AI.app' "$stage/app"
/usr/bin/ditto 'TrioBridge AI.service' "$stage/service"
/usr/bin/codesign --verify --deep --strict "$stage/app"
/usr/bin/codesign --verify --strict "$stage/service"
moved_app=0; moved_service=0; installed_app=0; installed_service=0
rollback() {
  [ "$installed_service" = 0 ] || /bin/rm -rf "$provider"
  [ "$installed_app" = 0 ] || /bin/rm -rf "$app"
  [ "$moved_app" = 0 ] || mv "$stage/old-app" "$app"
  [ "$moved_service" = 0 ] || mv "$stage/old-service" "$provider"
}
trap 'rollback' ERR
if [ -e "$app" ]; then mv "$app" "$stage/old-app"; moved_app=1; fi
if [ -e "$provider" ]; then mv "$provider" "$stage/old-service"; moved_service=1; fi
mv "$stage/app" "$app"; installed_app=1
mv "$stage/service" "$provider"; installed_service=1
lsregister='/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister'
"$lsregister" -f "$app"
"$lsregister" -f "$provider"
/System/Library/CoreServices/pbs -update
trap - ERR
echo '安装完成。在 Word、Excel 或 PowerPoint 的应用菜单中选择“服务 → 打开三桥 AI”。'
echo '如需 Office 内嵌窗格，请另行运行 Enable-Embedded-Taskpane.command；该入口需要本机 HTTPS 证书和联网加载微软 Office.js。'
