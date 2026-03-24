"""올리브청소년방과후 센터 활동앱 – Flask 메인"""
import os
import logging
from dotenv import load_dotenv

# Config 클래스 정의(import) 전에 반드시 load_dotenv() 호출
# → os.environ.get() 이 .env 값을 읽을 수 있도록 환경변수를 먼저 채움
load_dotenv()

from flask import Flask, render_template, request, jsonify, send_file, make_response
from config import Config

app = Flask(__name__)
app.config.from_object(Config)

# 오류 확인용: 앱 로그를 server_err.log에 추가
try:
    log_path = os.path.join(os.path.dirname(__file__), 'server_err.log')
    file_handler = logging.FileHandler(log_path, encoding='utf-8')
    file_handler.setLevel(logging.WARNING)
    app.logger.addHandler(file_handler)
except Exception:
    pass


@app.after_request
def add_cors_headers(response):
    """CORS 허용 – 같은 Wi‑Fi 핸드폰 접속·로컬 개발 시 fetch 오류 방지."""
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    return response


@app.before_request
def handle_options():
    if request.method == 'OPTIONS':
        r = make_response('', 204)
        r.headers['Access-Control-Allow-Origin'] = '*'
        r.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
        r.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        return r

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

from routes.daily_log import daily_log_bp
from routes.purchase_doc import purchase_doc_bp
from routes.plan import plan_bp
from services.hwp_service import get_photo_insert_log_path

app.register_blueprint(daily_log_bp)
app.register_blueprint(purchase_doc_bp)
app.register_blueprint(plan_bp)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/open-file', methods=['POST'])
def open_file():
    """생성된 HWP 파일을 로컬에서 직접 엽니다."""
    try:
        path = request.json.get('path', '')
        if not path or not os.path.exists(path):
            return jsonify(success=False, error="파일을 찾을 수 없습니다.")
        os.startfile(path)
        return jsonify(success=True)
    except Exception as e:
        return jsonify(success=False, error=str(e))


@app.route('/download')
def download_file():
    """생성된 HWP 파일을 브라우저로 직접 내려받습니다."""
    path = request.args.get('path', '')
    if not path:
        return "path 쿼리 파라미터가 필요합니다.", 400

    abs_path = os.path.abspath(path)
    base = os.path.abspath(app.config['OUTPUT_FOLDER'])
    if not abs_path.startswith(base) or not os.path.exists(abs_path):
        return "파일을 찾을 수 없습니다.", 404

    filename = os.path.basename(abs_path)
    return send_file(abs_path, as_attachment=True, download_name=filename)


@app.route('/photo-insert-status', methods=['POST'])
def photo_insert_status():
    """사진 삽입 워커 로그를 읽어 현재 상태를 반환합니다."""
    try:
        output_path = request.json.get('path', '')
        if not output_path:
            return jsonify(success=False, error="문서 경로가 필요합니다.")

        log_path = get_photo_insert_log_path(output_path)
        if not os.path.exists(log_path):
            return jsonify(
                success=True,
                status="not_started",
                log_path=log_path,
                log_text="로그 파일이 아직 생성되지 않았습니다."
            )

        with open(log_path, 'r', encoding='utf-8', errors='replace') as f:
            log_text = f.read()

        status = "running"
        if "STATUS:SUCCESS" in log_text:
            status = "success"
        elif "STATUS:ERROR" in log_text:
            status = "error"
        elif "STATUS:STARTED" in log_text:
            status = "running"

        return jsonify(
            success=True,
            status=status,
            log_path=log_path,
            log_text=log_text[-2000:]
        )
    except Exception as e:
        return jsonify(success=False, error=str(e))


def _get_local_ip():
    """같은 Wi-Fi에서 핸드폰 접속용 로컬 IP를 얻습니다."""
    try:
        import socket
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.settimeout(0)
        s.connect(("10.255.255.255", 1))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return None


if __name__ == '__main__':
    local_ip = _get_local_ip()
    print("=" * 55)
    print("  올리브청소년방과후 센터 활동앱")
    print("  PC: http://localhost:5006")
    if local_ip:
        print(f"  핸드폰(같은 Wi-Fi): http://{local_ip}:5006")
    print("=" * 55)
    app.run(debug=False, port=5006, host='0.0.0.0', threaded=False)
