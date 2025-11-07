#!/bin/bash

# LaunchAgent インストールスクリプト
# macOS用

echo "==================================="
echo "LaunchAgent インストール"
echo "==================================="

# 現在のユーザー名を取得
USERNAME=$(whoami)

# 現在のディレクトリの絶対パスを取得
SCRIPT_DIR=$(cd $(dirname $0); pwd)

# Pythonのパスを取得
PYTHON_PATH=$(which python3)

# plistファイルの作成
PLIST_FILE="$HOME/Library/LaunchAgents/com.mailautomation.plist"

echo "plistファイルを作成中..."

cat > "$PLIST_FILE" << EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.mailautomation</string>
    <key>ProgramArguments</key>
    <array>
        <string>${PYTHON_PATH}</string>
        <string>${SCRIPT_DIR}/main.py</string>
    </array>
    <key>WorkingDirectory</key>
    <string>${SCRIPT_DIR}</string>
    <key>StartInterval</key>
    <integer>600</integer>
    <key>RunAtLoad</key>
    <true/>
    <key>StandardOutPath</key>
    <string>${SCRIPT_DIR}/logs/stdout.log</string>
    <key>StandardErrorPath</key>
    <string>${SCRIPT_DIR}/logs/stderr.log</string>
    <key>EnvironmentVariables</key>
    <dict>
        <key>PATH</key>
        <string>/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin</string>
    </dict>
</dict>
</plist>
EOF

echo "plistファイルを作成しました: $PLIST_FILE"

# LaunchAgentのロード
echo "LaunchAgentをロード中..."
launchctl unload "$PLIST_FILE" 2>/dev/null
launchctl load "$PLIST_FILE"

echo ""
echo "==================================="
echo "インストール完了！"
echo "==================================="
echo ""
echo "ステータス確認:"
echo "launchctl list | grep mailautomation"
echo ""
echo "手動実行:"
echo "launchctl start com.mailautomation"
echo ""
echo "停止:"
echo "launchctl unload ~/Library/LaunchAgents/com.mailautomation.plist"
echo ""

