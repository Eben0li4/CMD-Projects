@echo off
code "C:\cmd\SSRS(KMRC)"
timeout /t 2 >nul
code -r -g "C:\cmd\SSRS(KMRC)\UpdateCrushingDump.py"
timeout /t 2 >nul
code --new-window
