#!/bin/sh

# Terminal launcher for SQL Console → Excel.
APP_FILE="/Users/user/test-agent/sql-console-to-excel/index.html"

if [ ! -f "$APP_FILE" ]; then
  printf '找不到 SQL Console → Excel：%s\n' "$APP_FILE" >&2
  printf '如果移动了项目，请重新运行 install-command.sh。\n' >&2
  exit 1
fi

case "$(uname -s)" in
  Darwin)
    open "$APP_FILE"
    ;;
  Linux)
    if command -v xdg-open >/dev/null 2>&1; then
      xdg-open "$APP_FILE" >/dev/null 2>&1 &
    else
      printf '无法打开浏览器：系统中没有 xdg-open。\n' >&2
      exit 1
    fi
    ;;
  *)
    printf '暂不支持当前系统：%s\n' "$(uname -s)" >&2
    exit 1
    ;;
esac
