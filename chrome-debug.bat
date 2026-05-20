@echo off
REM Запуск Chrome для рецептов в режиме отладки.
REM Двойной клик по этому файлу = запуск отладочного Chrome.
REM
REM Если у тебя Chrome установлен НЕ в "C:\Program Files\Google\Chrome\",
REM а в "C:\Program Files (x86)\Google\Chrome\" — поправь путь ниже.

start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" ^
  --remote-debugging-port=9222 ^
  --user-data-dir="%USERPROFILE%\chrome-debug-profile"
