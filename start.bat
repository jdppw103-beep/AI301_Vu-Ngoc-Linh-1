@echo off
chcp 65001 >nul
echo =======================================================
echo   DataLens AI - He thong phan tich du lieu khach hang
echo =======================================================
echo   Dia chi:  http://datalens.localhost:5000
echo   Giu cua so nay mo de he thong tiep tuc hoat dong.
echo =======================================================
start http://datalens.localhost:5000
dotnet run --urls "http://localhost:5000"
pause
