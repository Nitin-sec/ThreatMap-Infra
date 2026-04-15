#!/usr/bin/env bash
set -e

echo "====================================="
echo " ThreatMap Infra Installer"
echo "====================================="

# Install system dependencies
echo "[*] Installing system dependencies..."
sudo apt update
sudo apt install -y python3 python3-venv python3-pip \
    nmap nikto gobuster curl sslscan libreoffice

# Create hidden virtual environment
echo "[*] Setting up environment..."
python3 -m venv .venv

# Activate and install Python deps
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt

echo "[✔] Installation complete!"
echo "Run the tool using: ./ThreatMap-Infra"
