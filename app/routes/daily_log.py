"""일일활동일지 라우트"""
import os
import traceback
from flask import Blueprint, render_template, request, jsonify, current_app
from werkzeug.utils import secure_filename
from services import openai_service, hwp_service

daily_log_bp = Blueprint('daily_log', __name__)


def _user_friendly_error(e: Exception) -> str:
    """사용자에게 보여줄 에러 메시지로 변환 (한글/HWP 관련 자주 나는 오류 안내)."""
    msg = str(e).strip()
    if not msg:
        return "일시적인 오류가 발생했습니다. 다시 시도해 주세요."
    if "한글" in msg or "HWP" in msg or "연결" in msg or "GetActiveObject" in msg or "작업을 사용할 수 없습니다" in msg:
        return "한글 프로그램에 연결할 수 없습니다. 한글을 먼저 실행한 뒤 다시 시도해 주세요."
    if "hwp-mcp" in msg or "모듈" in msg:
        return "HWP 연동 모듈을 찾을 수 없습니다. hwp-mcp 경로를 확인해 주세요."
    if len(msg) > 200:
        return msg[:200] + "..."
    return msg

ALLOWED_EXT = {'png', 'jpg', 'jpeg', 'gif', 'webp'}


def _allowed(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXT


@daily_log_bp.route('/daily-log')
def daily_log_page():
    return render_template('daily_log.html')


@daily_log_bp.route('/daily-log/generate', methods=['POST'])
def generate_daily_log():
    try:
        date          = request.form.get('date', '')
        student_names = request.form.get('student_names', '').strip()

        # 사진별 메타데이터 수집 (최대 3장)
        photo_metas = []
        for i in range(1, 4):
            photo_metas.append({
                'time':  request.form.get(f'photo_time_{i}', '').strip(),
                'place': request.form.get(f'photo_place_{i}', '').strip(),
                'note':  request.form.get(f'photo_note_{i}', '').strip(),
            })

        # 사진 저장
        files = request.files.getlist('photos')
        if not files or all(f.filename == '' for f in files):
            return jsonify(success=False, error="사진을 1장 이상 업로드해 주세요.")

        saved_paths = []
        upload_dir = current_app.config['UPLOAD_FOLDER']
        for f in files:
            if f and _allowed(f.filename):
                fname = secure_filename(f.filename)
                path = os.path.join(upload_dir, fname)
                f.save(path)
                saved_paths.append(path)

        if not saved_paths:
            return jsonify(success=False, error="지원하는 이미지 형식(jpg, png, gif, webp)이 아닙니다.")

        # 실제 업로드된 사진 수에 맞게 메타 자르기
        photo_metas = photo_metas[:len(saved_paths)]

        # OpenAI로 활동기록 생성 (사진별 메타 + 아동 명단 포함)
        api_key = current_app.config.get('OPENAI_API_KEY', '')
        activities = openai_service.generate_activity_content(
            saved_paths, photo_metas, api_key, student_names=student_names
        )

        # HWP 생성
        data = {
            'date':             date,
            'student_names':    student_names,
            'time':             ', '.join(m['time']  for m in photo_metas if m['time']),
            'place':            ', '.join(m['place'] for m in photo_metas if m['place']),
            'special_note':     '\n'.join(m['note']  for m in photo_metas if m['note']),
            'activity_content': activities.get('full_text', ''),
            'activities':       activities,
            'photo_paths':      saved_paths,
        }
        output_path = hwp_service.create_daily_log(data, current_app.config['OUTPUT_FOLDER'])

        is_demo = not api_key or activities.get('_demo', False)
        return jsonify(
            success=True,
            activity_content=activities.get('full_text', ''),
            output_path=output_path,
            is_demo=is_demo,
            message=f"일일활동일지가 생성되었습니다.\n저장 위치: {output_path}"
        )

    except Exception as e:
        current_app.logger.warning("daily_log generate error: %s\n%s", e, traceback.format_exc())
        return jsonify(success=False, error=_user_friendly_error(e))
