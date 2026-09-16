@echo off
py -3.12 "%~dp0caption_keeper.py"
if errorlevel 1 (
  echo.
  echo 無法啟動 Caption Keeper。請確認 Python 3.12 與需求套件已安裝。
  pause
)

