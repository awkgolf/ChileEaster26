#!/bin/bash

echo "⛏️ GEOLOGICAL PROJECT: TABLET INITIALIZATION"
echo "------------------------------------------"

# 1. Update system packages
echo "🔄 Updating system packages..."
pkg update -y && pkg upgrade -y

# 2. Install Node.js
echo "📦 Installing Node.js..."
pkg install nodejs -y

# 3. Request Storage Access (Crucial for Android)
echo "📂 Requesting storage access..."
termux-setup-storage

# 4. Install Project Dependencies
echo "🏗️ Building local node_modules..."
npm install

echo "------------------------------------------"
echo "✅ SETUP COMPLETE."
echo "💡 To build your journal, type: node index.js"