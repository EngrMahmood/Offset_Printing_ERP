@echo off
REM ============================================================
REM  Offset ERP - Live server log viewer.
REM  Purely a visibility window; does NOT run the server itself
REM  (that's the headless "Offset ERP Server" boot task). Opens on
REM  login and tails backups\server.log so you can see it's alive.
REM ============================================================
title Offset ERP - Server Log (live)
powershell -NoExit -Command "Get-Content -Path 'E:\Offset_Printing_ERP\backups\server.log' -Wait -Tail 50"
