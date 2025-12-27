@echo off
echo ========================================
echo      Installing Python libraries...
echo ========================================
echo.
python -m pip install pandas xlrd xlwt openpyxl -i https://pypi.tuna.tsinghua.edu.cn/simple
echo.
echo Done!
echo.
pause
