@echo off
chcp 65001 >nul
cd /d "%~dp0..\.."
python -u -m RFQ.tags_rfp_compare.compare_step4_files "%~1"
pause
