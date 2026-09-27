#!/bin/bash
set -euo pipefail

echo "⛏️ GEOLOGICAL PROJECT: TABLET INITIALIZATION"
echo "------------------------------------------"

echo "🔄 Updating system packages..."
pkg update -y && pkg upgrade -y

echo "📦 Installing Node.js..."
pkg install nodejs -y

echo "📂 Requesting storage access..."
termux-setup-storage

echo "🏗️ Installing project dependencies..."
npm install

echo "✅ SETUP COMPLETE."
echo "💡 To validate data: npm run validate"
echo "💡 To build your journal: npm run build"
