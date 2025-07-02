#!/bin/bash

echo "🚀 Starting Excel Taak-Norm Verrijker Apps..."

# Add local bin to path
export PATH=$PATH:/home/ubuntu/.local/bin

# Start all apps in background
echo "📋 Starting Launcher (Choice Menu) on port 8500..."
streamlit run launcher.py --server.headless true --server.port 8500 &

echo "🤖 Starting Automatic App on port 8501..."
streamlit run app.py --server.headless true --server.port 8501 &

echo "🔧 Starting Manual App on port 8502..."
streamlit run app_manual.py --server.headless true --server.port 8502 &

echo ""
echo "⏳ Waiting for apps to start..."
sleep 10

echo ""
echo "=== STATUS CHECK ==="
echo ""

echo "🚀 Launcher (Keuze Menu):"
if curl -s http://localhost:8500 > /dev/null; then
    echo "✅ Online: http://localhost:8500"
else
    echo "❌ Offline"
fi

echo ""
echo "🤖 Automatische App:"
if curl -s http://localhost:8501 > /dev/null; then
    echo "✅ Online: http://localhost:8501"
else
    echo "❌ Offline"
fi

echo ""
echo "🔧 Handmatige App:"
if curl -s http://localhost:8502 > /dev/null; then
    echo "✅ Online: http://localhost:8502"
else
    echo "❌ Offline"
fi

echo ""
echo "🎯 READY TO USE!"
echo ""
echo "📖 Instructies:"
echo "1. Ga naar http://localhost:8500 voor het keuze menu"
echo "2. Of ga direct naar:"
echo "   • http://localhost:8501 (automatische versie)"
echo "   • http://localhost:8502 (handmatige versie - AANBEVOLEN)"
echo ""
echo "💡 TIP: Gebruik de handmatige versie als je Excel lees-fouten krijgt!"
echo ""
echo "🛑 Stop alle apps: pkill -f streamlit"