#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"
streamlit run STREAMFINAL.py \
  --server.port="${PORT:-10000}" \
  --server.address=0.0.0.0 \
  --server.headless=true \
  --server.fileWatcherType=none \
  --browser.gatherUsageStats=false
