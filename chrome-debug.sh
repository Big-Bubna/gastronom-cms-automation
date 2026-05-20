#!/bin/bash
# Запуск Chrome для рецептов в режиме отладки.
# Сначала сделай исполняемым: chmod +x chrome-debug.sh
# Затем запускай: ./chrome-debug.sh

# Полностью закрой Chrome перед запуском (⌘+Q, не просто закрытие окна)

/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome \
  --remote-debugging-port=9222 \
  --user-data-dir="$HOME/chrome-debug-profile"
