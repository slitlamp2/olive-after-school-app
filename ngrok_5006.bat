@echo off
chcp 65001 >nul
echo ========================================
echo  방과후센터앱 외부 접속용 ngrok (포트 5006)
echo ========================================
echo.
echo [필수] Flask 앱이 먼저 실행 중이어야 합니다.
echo        (예: app 폴더에서 python app.py → localhost:5006)
echo.
echo [처음 1회] ngrok 대시보드에서 Authtoken 복사 후 아래 실행:
echo        ngrok config add-authtoken 여기에_토큰_붙여넣기
echo        https://dashboard.ngrok.com/get-started/your-authtoken
echo.
echo 잠시 후 나오는 Forwarding 주소(https://....ngrok-free.app)를
echo 휴대폰/다른 PC 브라우저에 입력하세요.
echo ========================================
echo.
ngrok http 5006
pause
