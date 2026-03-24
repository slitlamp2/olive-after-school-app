@echo off
REM 항상 배치파일이 있는 app 폴더 기준으로 실행
cd /d "%~dp0"

chcp 65001 > nul
echo.
echo  ==========================================
echo   올리브청소년방과후 센터 활동앱
echo  ==========================================
echo.
echo  필요한 패키지를 확인합니다...
pip install flask openai python-dotenv pywin32 --quiet

echo.
echo  앱을 시작합니다...
echo  브라우저에서 http://localhost:5000 을 열어 주세요.
echo.
python app.py

pause
