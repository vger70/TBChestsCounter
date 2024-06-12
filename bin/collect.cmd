@echo off

cd c:\S\TB\TB_Chests-main\TB\bin

:start
py .\tb.py --config ..\config\Roma.cfg --capture
timeout 120
goto :start