#!/bin/sh

set -eu

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
SOURCE_FILE="$SCRIPT_DIR/sql"
INSTALL_DIR="$HOME/.local/bin"
TARGET_FILE="$INSTALL_DIR/sql"

mkdir -p "$INSTALL_DIR"

# Bake in the current project path so the command works from every directory.
sed "s|^APP_FILE=.*$|APP_FILE=\"$SCRIPT_DIR/index.html\"|" "$SOURCE_FILE" > "$TARGET_FILE"
chmod +x "$TARGET_FILE"

printf '安装成功：%s\n' "$TARGET_FILE"
printf '现在可以在终端的任意目录输入 sql 打开工具。\n'
