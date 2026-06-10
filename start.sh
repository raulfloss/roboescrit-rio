#!/bin/bash
echo "[start.sh] Iniciando Xvfb no display :99..."
Xvfb :99 -screen 0 1280x720x24 -ac &
XVFB_PID=$!
sleep 2

if kill -0 $XVFB_PID 2>/dev/null; then
    export DISPLAY=:99
    echo "[start.sh] Xvfb OK | DISPLAY=$DISPLAY"
else
    echo "[start.sh] Xvfb falhou — Chrome rodará sem display (headless)"
fi

echo "[start.sh] Iniciando servidor Flask..."
exec python server.py
