"""계획서 라우트"""
import os
import traceback
from flask import Blueprint, render_template, request, jsonify, current_app, session
from werkzeug.utils import secure_filename
from services import openai_service, hwp_service
from datetime import date as dt_date

plan_bp = Blueprint('plan', __name__)
ALLOWED_EXT = {'png', 'jpg', 'jpeg', 'gif', 'webp'}


def _user_friendly_error(e: Exception) -> str:
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


def _allowed(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXT


@plan_bp.route('/plan')
def plan_page():
    purchase_summary = session.get('purchase_summary', '')
    return render_template('plan.html', purchase_summary=purchase_summary)


@plan_bp.route('/plan/generate', methods=['POST'])
def generate_plan():
    try:
        purchase_summary = request.form.get('purchase_summary', '')
        plan_date = request.form.get('plan_date', str(dt_date.today()))

        # 품위서 단계에서 저장해 둔 품목/합계 정보를 세션에서 가져와
        # 계획서 HWP 표에도 그대로 사용할 수 있게 전달
        purchase_items = session.get('purchase_items', [])
        purchase_total_amount = session.get('purchase_total_amount', '')

        files = request.files.getlist('photos')
        saved_paths = []
        upload_dir = current_app.config['UPLOAD_FOLDER']
        for f in files:
            if f and f.filename and _allowed(f.filename):
                fname = secure_filename(f.filename)
                path = os.path.join(upload_dir, fname)
                f.save(path)
                saved_paths.append(path)

        api_key = current_app.config.get('OPENAI_API_KEY', '')
        plan_content = openai_service.generate_plan_content(
            saved_paths, purchase_summary, api_key
        )

        data = {
            'plan_date': plan_date,
            'purchase_summary': purchase_summary,
            'purchase_items': purchase_items,
            'purchase_total_amount': purchase_total_amount,
            **plan_content,
        }
        output_path = hwp_service.create_plan(
            data,
            current_app.config['OUTPUT_FOLDER'],
            image_paths=saved_paths,
        )

        is_demo = plan_content.get('_demo', False)
        return jsonify(
            success=True,
            plan_content=plan_content,
            output_path=output_path,
            is_demo=is_demo,
            message=f"계획서가 생성되었습니다.\n저장 위치: {output_path}"
        )

    except Exception as e:
        current_app.logger.warning("plan generate error: %s\n%s", e, traceback.format_exc())
        return jsonify(success=False, error=_user_friendly_error(e))
