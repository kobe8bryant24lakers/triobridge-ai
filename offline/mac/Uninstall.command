#!/bin/bash
set -euo pipefail
if /usr/bin/pgrep -x 'TrioBridge AI' >/dev/null; then echo '请先关闭三桥 AI 工作台，再双击卸载。'; exit 1; fi
remove_owned() {
  if [ -d "$1" ] && [ "$(/usr/libexec/PlistBuddy -c 'Print CFBundleIdentifier' "$1/Contents/Info.plist")" = "$2" ]; then
    /System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister -u "$1" || true
    /bin/rm -rf "$1"
  fi
}
remove_owned "$HOME/Library/Services/TrioBridge AI.service" ai.triobridge.workbench.service
remove_owned "$HOME/Applications/TrioBridge AI.app" ai.triobridge.workbench
/System/Library/CoreServices/pbs -update
echo '已移除三桥 AI 工作台和 Office 入口。模型设置和凭据保留在当前账户中。'
