"""품위서 라우트"""
import os
import traceback
import uuid
from flask import Blueprint, render_template, request, jsonify, current_app, session
from werkzeug.utils import secure_filename
from services import openai_service, hwp_service

purchase_doc_bp = Blueprint('purchase_doc', __name__)
ALLOWED_EXT = {'png', 'jpg', 'jpeg', 'gif', 'webp'}
# 영수증은 장마다 OpenAI 1회 호출 — 입력 슬롯 최대 3개
MAX_RECEIPT_IMAGES = 3


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


@purchase_doc_bp.route('/purchase-doc')
def purchase_doc_page():
    return render_template('purchase_doc.html')


@purchase_doc_bp.route('/purchase-doc/generate', methods=['POST'])
def generate_purchase_doc():
    try:
        slot_keys = ("receipt1", "receipt2", "receipt3")
        to_save = []
        for key in slot_keys:
            f = request.files.get(key)
            if f and f.filename and f.filename.strip() and _allowed(f.filename):
                to_save.append(f)

        if not to_save:
            return jsonify(success=False, error="영수증 사진을 1장 이상 업로드해 주세요.")

        if len(to_save) > MAX_RECEIPT_IMAGES:
            return jsonify(
                success=False,
                error=f"영수증 사진은 최대 {MAX_RECEIPT_IMAGES}장까지 업로드할 수 있습니다.",
            )

        saved_paths = []
        upload_dir = current_app.config['UPLOAD_FOLDER']
        for f in to_save:
            base = secure_filename(f.filename) or "receipt.jpg"
            fname = f"{uuid.uuid4().hex[:12]}_{base}"
            path = os.path.join(upload_dir, fname)
            f.save(path)
            saved_paths.append(path)

        if not saved_paths:
            return jsonify(
                success=False,
                error="지원하는 이미지 형식(jpg, png, gif, webp)의 영수증을 업로드해 주세요.",
            )

        api_key = current_app.config.get('OPENAI_API_KEY', '')
        receipt_data = openai_service.extract_receipt_data(saved_paths, api_key)

        output_path = hwp_service.create_purchase_doc(
            receipt_data,
            current_app.config['OUTPUT_FOLDER'],
            image_paths=saved_paths,
        )

        # 세션에 품위서 데이터 저장 (계획서에서 사용)
        items_summary = "\n".join(
            f"- {it['name']} {it['qty']}{it['unit']} / {it['amount']}원"
            for it in receipt_data.get('items', [])
        )
        session['purchase_summary'] = (
            f"구입일자: {receipt_data.get('purchase_date','')}\n"
            f"매장: {receipt_data.get('store_name','')}\n"
            f"품목:\n{items_summary}\n"
            f"합계: {receipt_data.get('total_amount','')}원"
        )

        # 계획서에서 품위서 표를 그대로 사용할 수 있도록
        # 품목 리스트와 합계 금액도 세션에 그대로 저장
        session['purchase_items'] = receipt_data.get('items', [])
        session['purchase_total_amount'] = receipt_data.get('total_amount', '')

        is_demo = receipt_data.get('_demo', False)
        return jsonify(
            success=True,
            receipt_data=receipt_data,
            output_path=output_path,
            is_demo=is_demo,
            message=f"품위서가 생성되었습니다.\n저장 위치: {output_path}"
        )

    except Exception as e:
        current_app.logger.warning("purchase_doc generate error: %s\n%s", e, traceback.format_exc())
        return jsonify(success=False, error=_user_friendly_error(e))
