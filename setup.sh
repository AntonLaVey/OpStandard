#!/bin/bash

# Raspberry Pi Standards Viewer - GitHub-Based Setup
# Repository: https://github.com/AntonLaVey/OpStandard

if [ "$(id -u)" -ne 0 ]; then
  echo "Please run this script with sudo: sudo ./setup.sh"
  exit 1
fi

GITHUB_USER="AntonLaVey"
GITHUB_REPO="OpStandard"
GITHUB_BRANCH="main"
PYTHON_FILE_URL="https://raw.githubusercontent.com/$GITHUB_USER/$GITHUB_REPO/$GITHUB_BRANCH/image_viewer.py"

echo "========================================="
echo "Raspberry Pi Standards Viewer Setup"
echo "========================================="
echo ""

REAL_USER="${SUDO_USER:-$(whoami)}"
HOME_DIR=$(getent passwd "$REAL_USER" | cut -d: -f6)
APP_DIR="$HOME_DIR/pi_photo_app"
APP_SCRIPT_PATH="$APP_DIR/image_viewer.py"
LOG_DIR="/var/log/pi-photo-viewer"
LOG_FILE="$LOG_DIR/app.log"

echo "Installing for user: $REAL_USER"
echo ""

echo "[1/7] Installing system packages..."
apt-get update
apt-get install -y python3-tk python3-pil.imagetk inotify-tools libreoffice poppler-utils imagemagick curl

if [ $? -ne 0 ]; then
    echo "Error: Failed to install packages."
    exit 1
fi

echo "[2/7] Installing Python packages..."
# Download requirements.txt from GitHub
REQUIREMENTS_URL="https://raw.githubusercontent.com/$GITHUB_USER/$GITHUB_REPO/$GITHUB_BRANCH/requirements.txt"
curl -fsSL "$REQUIREMENTS_URL" -o /tmp/requirements.txt 2>/dev/null || true

if [ -f /tmp/requirements.txt ]; then
    pip3 install -r /tmp/requirements.txt --break-system-packages 2>/dev/null || pip3 install -r /tmp/requirements.txt
    rm /tmp/requirements.txt
else
    # Fallback to direct installation if requirements.txt not available
    pip3 install openpyxl --break-system-packages 2>/dev/null || pip3 install openpyxl
fi

echo "[3/7] Setting up logging and cache..."
mkdir -p "$LOG_DIR"
chown "$REAL_USER":"$REAL_USER" "$LOG_DIR"
chmod 755 "$LOG_DIR"
touch "$LOG_FILE"
chown "$REAL_USER":"$REAL_USER" "$LOG_FILE"

# Create cache directory with secure permissions
CACHE_DIR="/var/cache/pi-photo-viewer"
mkdir -p "$CACHE_DIR"
chown "$REAL_USER":"$REAL_USER" "$CACHE_DIR"
chmod 700 "$CACHE_DIR"  # Only owner can access

cat > /etc/logrotate.d/pi-photo-viewer << EOF
/var/log/pi-photo-viewer/app.log {
    daily
    rotate 7
    compress
    delaycompress
    notifempty
    create 0644 $REAL_USER $REAL_USER
}
EOF

echo "[4/7] Creating application directory..."
mkdir -p "$APP_DIR"
chown "$REAL_USER":"$REAL_USER" "$APP_DIR"

echo "[5/7] Downloading Python application from GitHub..."
echo "URL: $PYTHON_FILE_URL"
curl -fsSL "$PYTHON_FILE_URL" -o "$APP_SCRIPT_PATH"

if [ $? -ne 0 ]; then
    echo "Error: Failed to download Python file from GitHub."
    echo "Please check:"
    echo "  1. Your GitHub username is correct (AntonLaVey)"
    echo "  2. The repository exists and is public (OpStandard)"
    echo "  3. The file 'image_viewer.py' exists in the repo"
    exit 1
fi

chown "$REAL_USER":"$REAL_USER" "$APP_SCRIPT_PATH"
chmod +x "$APP_SCRIPT_PATH"

echo "[6/7] Configuring systemd services..."

# LibreOffice persistent listener - DISABLED
# We'll use fresh processes instead to avoid hangs
# cat > /etc/systemd/system/libreoffice-listener.service << 'EOF'
# [Unit]
# Description=LibreOffice Headless Listener
# Before=pi-photo-viewer.service
#
# [Service]
# Type=simple
# User=pi
# ExecStart=/usr/bin/libreoffice --headless --accept="socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" --nofirststartwizard --nologo --norestore --disable-extension-update
# Restart=always
# RestartSec=10
# TimeoutStopSec=15
# KillMode=mixed
#
# [Install]
# WantedBy=multi-user.target
# EOF

# Main application service
cat > /etc/systemd/system/pi-photo-viewer.service << EOF
[Unit]
Description=Raspberry Pi Standards Viewer
After=graphical.target

[Service]
Type=simple
User=$REAL_USER
Environment="DISPLAY=:0"
Environment="XAUTHORITY=$HOME_DIR/.Xauthority"
ExecStartPre=/bin/sleep 5
ExecStart=/usr/bin/python3 $APP_SCRIPT_PATH
ExecStopPost=/bin/pkill -9 -f "soffice|libreoffice"
Restart=on-failure
RestartSec=10
StandardOutput=journal
StandardError=journal
TimeoutStopSec=10
KillMode=mixed
KillSignal=SIGTERM

[Install]
WantedBy=graphical.target
EOF

# Desktop shortcut
mkdir -p "$HOME_DIR/Desktop"
cat > "$HOME_DIR/Desktop/PhotoViewer.desktop" << EOF
[Desktop Entry]
Version=1.0
Name=Pi Standards Viewer
Comment=Excel and Image Standards Display
Exec=python3 $APP_SCRIPT_PATH
Icon=folder-pictures
Terminal=false
Type=Application
Categories=Graphics;
EOF

chown "$REAL_USER":"$REAL_USER" "$HOME_DIR/Desktop/PhotoViewer.desktop"
chmod +x "$HOME_DIR/Desktop/PhotoViewer.desktop"

usermod -a -G video "$REAL_USER"
chmod 644 /etc/systemd/system/pi-photo-viewer.service

echo "[7/7] Disabling USB pop-ups..."
SYSTEM_CONFIG_FILE="/etc/xdg/pcmanfm/LXDE-pi/pcmanfm.conf"
if [ -f "$SYSTEM_CONFIG_FILE" ]; then
    if ! grep -q "\[volume\]" "$SYSTEM_CONFIG_FILE"; then
        echo -e "\n[volume]" >> "$SYSTEM_CONFIG_FILE"
    fi
    if sed -n '/\[volume\]/,/\[/p' "$SYSTEM_CONFIG_FILE" | grep -q 'autorun='; then
        sed -i '/\[volume\]/,/\[/s/autorun=.*/autorun=0/' "$SYSTEM_CONFIG_FILE"
    else
        sed -i '/\[volume\]/a autorun=0' "$SYSTEM_CONFIG_FILE"
    fi
fi

echo ""
echo "========================================="
echo "✅ Installation Complete!"
echo "========================================="
echo ""
echo "Application downloaded from GitHub:"
echo "  $PYTHON_FILE_URL"
echo ""
echo "To update in the future, run:"
echo "  curl -fsSL $PYTHON_FILE_URL -o $APP_SCRIPT_PATH"
echo "  sudo systemctl restart pi-photo-viewer"
echo ""
echo "Or use the quick update command:"
echo "  update-viewer"
echo ""
echo "Commands:"
echo "  Status: sudo systemctl status pi-photo-viewer"
echo "  Logs:   sudo journalctl -u pi-photo-viewer -f"
echo "  Update: curl -fsSL $PYTHON_FILE_URL -o $APP_SCRIPT_PATH && sudo systemctl restart pi-photo-viewer"
echo ""
echo "Please REBOOT now: sudo reboot"
echo "========================================="
