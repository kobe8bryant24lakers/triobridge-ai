#!/bin/bash
# This file is also sourced by the isolated installer tests.
set -euo pipefail

triobridge_manage() {
  local action="$1" user_root="$2" package_root="$3"
  local app directory destination actual stage
  case "$action" in install|remove) ;; *) echo 'Unknown action.' >&2; return 1 ;; esac
  if [ "$action" = install ]; then
    actual=$(shasum -a 256 "$package_root/manifest.xml")
    if [ "${actual%% *}" != '__MANIFEST_SHA256__' ]; then
      echo '安装包中的清单缺失或损坏，请重新下载安装包。' >&2
      return 1
    fi
  fi
  # Preflight every destination before changing any of them.
  for app in Word Excel Powerpoint; do
    directory="$user_root/Library/Containers/com.microsoft.$app/Data/Documents/wef"
    destination="$directory/triobridge-ai-online.xml"
    if [ -L "$directory" ] || [ -L "$destination" ]; then
      echo "拒绝修改符号链接：$destination" >&2
      return 1
    fi
    if [ -e "$destination" ] && ! grep -Fq '<Id>__HOSTED_MANIFEST_ID__</Id>' "$destination"; then
      echo "同名文件不属于三桥 AI 在线版，请先检查：$destination" >&2
      return 1
    fi
  done
  for app in Word Excel Powerpoint; do
    directory="$user_root/Library/Containers/com.microsoft.$app/Data/Documents/wef"
    destination="$directory/triobridge-ai-online.xml"
    if [ "$action" = install ]; then
      mkdir -p "$directory"
      stage=$(mktemp "$directory/.triobridge.XXXXXX")
      if ! cp "$package_root/manifest.xml" "$stage" || ! chmod 644 "$stage" || ! mv -f "$stage" "$destination"; then
        rm -f "$stage"
        return 1
      fi
      echo "已安装：$app"
    else
      rm -f "$destination"
      echo "已移除安装文件：$app"
    fi
  done
}

if [ "${BASH_SOURCE[0]}" = "$0" ]; then
  if [ "$(uname -s)" != Darwin ]; then
    echo '此安装包仅适用于 macOS。' >&2
    exit 1
  fi
  package_root=$(cd "$(dirname "$0")" && pwd)
  triobridge_manage "${1:-install}" "$HOME" "$package_root"
  echo '请保存文档并完全退出 Office，然后重新打开。'
  if [ "${1:-install}" = install ]; then
    echo '在「主页 → 加载项」选择「TrioBridge AI · 三桥 AI（在线版）」。'
    echo '然后在「模型设置」选择服务，填入 API Key，点击「连接并保存」。'
  else
    echo '若 Office 仍显示缓存入口，请按微软说明清理加载项缓存。'
    echo '模型配置和密钥未删除，如需清除请先在插件内删除。'
  fi
fi
