"""
OpenAI 서비스 – 사진을 분석하여 텍스트를 생성합니다.
urllib.request 직접 호출 방식으로 Python 버전 호환성 문제를 회피합니다.
API 키가 없으면 데모 텍스트를 반환합니다.
"""
import base64
import json
import os
import urllib.request
import urllib.error


OPENAI_API_URL = "https://api.openai.com/v1/chat/completions"


def _encode_image(image_path: str) -> str:
    with open(image_path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def _build_image_content(image_paths: list, detail: str = "low") -> list:
    """이미지 경로 목록을 OpenAI 메시지 content 배열로 변환합니다. detail: 'low' | 'high' (영수증은 'high' 권장)."""
    content = []
    for path in image_paths:
        ext = os.path.splitext(path)[1].lower().lstrip(".")
        mime = "image/jpeg" if ext in ("jpg", "jpeg") else f"image/{ext}"
        content.append({
            "type": "image_url",
            "image_url": {
                "url": f"data:{mime};base64,{_encode_image(path)}",
                "detail": detail
            }
        })
    return content


def _call_openai(api_key: str, messages: list, max_tokens: int = 800) -> str:
    """OpenAI Chat Completions API를 urllib로 직접 호출합니다."""
    payload = json.dumps({
        "model": "gpt-4o",
        "messages": messages,
        "max_tokens": max_tokens
    }).encode("utf-8")

    req = urllib.request.Request(
        OPENAI_API_URL,
        data=payload,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        },
        method="POST"
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        result = json.loads(resp.read().decode("utf-8"))
    return result["choices"][0]["message"]["content"].strip()


def _extract_json(raw: str) -> dict:
    """응답 문자열에서 JSON 블록을 최대한 유연하게 추출합니다."""
    import re

    # 1) 코드 블록(``` ```) 안에서 먼저 시도
    if "```" in raw:
        parts = raw.split("```")
        for part in parts[1::2]:
            if part.startswith("json"):
                part = part[4:]
            part = part.strip()
            if part:
                try:
                    return json.loads(part)
                except json.JSONDecodeError:
                    pass

    # 2) { … } 범위 추출 후 직접 파싱
    raw = raw.strip()
    start = raw.find('{')
    end   = raw.rfind('}')
    if start != -1 and end > start:
        candidate = raw[start:end + 1]
        try:
            return json.loads(candidate)
        except json.JSONDecodeError:
            # 3) JSON 값 내 제어문자·리터럴 줄바꿈 정리 후 재시도
            try:
                # 각 줄 내 리터럴 줄바꿈을 공백으로 대체
                fixed = re.sub(r'(?<!\\)\n', ' ', candidate)
                fixed = re.sub(r'(?<!\\)\r', ' ', fixed)
                return json.loads(fixed)
            except json.JSONDecodeError:
                pass

    # 4) 전체 원문 직접 시도
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return {}


def _rewrite_official_style_to_teacher(text: str) -> str:
    """
    OpenAI가 문장을 공문서 문체(~하였다, ~실시하였다 등)로 만들면,
    사용자 요청대로 따뜻한 교사(보호자에게 설명) 말투로 최대한 치환합니다.

    구조/의미를 크게 깨지 않기 위해 '자주 나오는 종결/표현'만 최소 규칙으로 변환합니다.
    """
    if not text:
        return text

    # 자주 등장하는 표현을 먼저 치환
    replacements = [
        ("실시하였다", "실시하였어요"),
        ("진행되었다", "진행하였어요"),
        ("진행하였다", "진행하였어요"),
        ("참여하였다", "참여하였어요"),
        ("참여하였음", "참여하였어요"),
        ("도와주었다", "도와주었어요"),
        ("제공하였다", "제공하였어요"),
        ("완성하였다", "완성하였어요"),
        ("완료되었습니다", "잘 마무리하였어요"),
        ("정리하였다", "정리하였어요"),
        ("정리하였다.", "정리하였어요."),
        ("인계되었다", "인계하였어요"),
        ("인계하였다", "인계하였어요"),
        ("설명하였다", "설명하였어요"),
        ("확인하였다", "확인하였어요"),
        ("기다리었다", "기다리었어요"),
        ("배우었다", "배우었어요"),
    ]
    for old, new in replacements:
        text = text.replace(old, new)

    # 공문서 종결 패턴을 완화 치환
    import re as _re
    # 이미 '하였음'으로 끝난 경우는 그대로 유지
    # (아래 변환들은 주로 '하였다' 형태를 보고서 종결로 정리하기 위함)
    text = _re.sub(r"하였다", "하였어요", text)
    text = _re.sub(r"되었다", "됐어요", text)
    text = _re.sub(r"되었", "됐", text)

    # 불필요한 공백 정리
    return " ".join(text.split())


def _rewrite_teacher_ending_to_eum(text: str) -> str:
    """
    사용자 요청용 후처리:
    문장 끝의 '~어요' 말투를 '~음' 문체로 변환합니다.

    예)
      '보냈어요.' -> '보냈음.'
      '했어요'    -> '했음'
    """
    if not text:
        return text

    import re as _re

    # 특별 처리: '주셨어요.' 같은 형태는 보고서 종결어미로 정리
    text = _re.sub(
        r"주셨어요(?=(?:\s*[\"'”]*\s*[.!?]|(?:\s*[\"'”]*)\s*$))",
        "주셨음",
        text,
    )

    # '하였어요/했어요/보였어요'는 어미만 정확히 보고서형으로 치환
    text = _re.sub(r"하였어요(?=(?:\s*[\"'”]*\s*[.!?]|(?:\s*[\"'”]*)\s*$))", "하였음", text)
    text = _re.sub(r"했어요(?=(?:\s*[\"'”]*\s*[.!?]|(?:\s*[\"'”]*)\s*$))", "했음", text)
    text = _re.sub(r"보였어요(?=(?:\s*[\"'”]*\s*[.!?]|(?:\s*[\"'”]*)\s*$))", "보였음", text)

    # 문장/줄 끝에서 '~어요'가 나오면 '~음'으로 치환
    # (뒤에 공백/따옴표/마침표가 붙는 경우까지 대응)
    text = _re.sub(
        r"([가-힣A-Za-z0-9]+)어요(?=(?:\s*[\"'”]*)\s*[.!?]|(?:\s*[\"'”]*)\s*$)",
        r"\1음",
        text,
    )

    # '하더라고요.', '말하더라고요.'처럼 '~라고요'로 끝나는 경우도 음체로 정리
    # (예: 하더라고요. -> 하더라고.)
    text = _re.sub(
        r"라고요(?=(?:[\"'”]+\s*)?[.!?]|$)",
        "라고",
        text,
    )

    # 사회복지사 보고서 문체로 마무리:
    # '...하더라고(요).' 패턴은 '...하였음'으로 바꿔줍니다.
    # 예) '연습하더라고요.' -> '연습하였음.'
    text = _re.sub(
        r"하더라고요(?=(?:\s*[\"'”]*)\s*[.!?]|(?:\s*[\"'”]*)\s*$)",
        "하였음",
        text,
    )
    text = _re.sub(
        r"하더라고(?=(?:\s*[\"'”]*)\s*[.!?]|(?:\s*[\"'”]*)\s*$)",
        "하였음",
        text,
    )

    # 중복 공백 정리
    return " ".join(text.split())


# ──────────────────────────────────────────────
# 일일활동일지 – 시간대별 활동 내용 생성
# ──────────────────────────────────────────────
def generate_activity_content(image_paths: list, photo_metas: list,
                               api_key: str, student_names: str = '') -> dict:
    """
    사진 3장 + 사진별 메타데이터를 바탕으로 시간대별 프로그램 참여 내용을 생성합니다.
    photo_metas: [{'time':..., 'program':..., 'place':..., 'note':...}, ...]  사진 순서에 맞게

    Returns dict:
        {
          "activity_1430": str,
          "activity_1500": str,
          "activity_1600": str,
          "activity_1700": str,
          "activity_1800": str,
          "special_note":  str,
          "full_text":     str,
        }
    """
    if not api_key:
        return _demo_activity(photo_metas)

    # 사진별 컨텍스트 문자열 구성
    photo_ctx_lines = []
    for i, meta in enumerate(photo_metas, 1):
        parts = [f"[사진 {i}]"]
        if meta.get('time'):  parts.append(f"활동시간: {meta['time']}")
        if meta.get('program'): parts.append(f"프로그램내용(직접입력): {meta['program']}")
        if meta.get('place'): parts.append(f"활동장소: {meta['place']}")
        if meta.get('note'):  parts.append(f"특이사항: {meta['note']}")
        photo_ctx_lines.append("  ".join(parts))
    photo_ctx = "\n".join(photo_ctx_lines) if photo_ctx_lines else "정보 없음"

    # 아동 명단 문자열
    names_ctx = f"참여 아동: {student_names}" if student_names else ""

    img_content = _build_image_content(image_paths)
    prompt_text = (
        "아래 사진들은 발달장애를 가진 청소년들의 방과후 프로그램 활동 장면입니다.\n\n"
        + (f"【참여 아동 명단】\n{names_ctx}\n\n" if names_ctx else "")
        +
        "【사진별 활동 정보】\n"
        f"{photo_ctx}\n\n"
        "사진에 보이는 아동들의 모습·표정·동작·사용 교구·재료·상호작용을 분석하여 "
        "아래 JSON 형식으로 작성해 주세요.\n\n"
        "- 각 사진에 '프로그램내용(직접입력)'이 있으면 해당 사진의 시간대·장면과 맞게 참여 서술에 반드시 반영하고, "
        "없으면 사진만으로 판단하세요.\n"
        "【문체 기준 – 매우 중요】\n"
        "- 사회복지사가 관찰 내용을 보고서에 작성하듯 기술합니다.\n"
        "- 문장 끝은 반드시 '~했음', '~하였음', '~보였음', '~전해 드렸음', '~제공하였음' 등 보고서 종결어미로만 마무리합니다.\n"
        "- '~어요', '~해요', '~하더라고요', '~라고요' 같은 구어/회고 종결어미는 절대 사용하지 않습니다.\n"
        "- 특정 종결이 어색하면 같은 의미로 보고서 종결(~했음/~하였음)로 바꿔 주세요.\n"
        "- '이 사진에서는', '사진에서'처럼 사진 자체를 직접 언급하지 말고, "
        "  아이들의 활동과 상황만 자연스럽게 서술하세요.\n"
        "- 모든 문장은 한국어로만 작성하고, gently, refresh 같은 영어 단어는 쓰지 말고 "
        "  '부드럽게', '재충전', '휴식'처럼 한국어 표현으로 바꿔 주세요.\n"
        "  예) '오늘 ○○이가 친구들이랑 같이 블록 쌓기를 신나게 즐겼음.'\n"
        "      '처음엔 조금 망설였지만 선생님이 옆에서 도와드리니 금방 따라왔음.'\n"
        "      '간식 시간엔 직접 만든 샌드위치를 맛있게 먹었음.'\n"
        + (
            f"- {student_names} 이름은 첫 문장에서만 한 번 사용하고,\n"
            "  이후엔 '아이가', '○○이가' 대신 '또', '그리고', '이어서' 등으로 자연스럽게 이어 주세요.\n"
            if student_names else
            "- 아이 이름 대신 '아이가', '친구들이랑', '다 같이' 같은 표현을 사용하세요.\n"
        )
        +
        "- 각 항목은 사진 속 실제 장면을 근거로 200~260자로 써 주세요.\n"
        "- 반드시 구체적으로 서술하세요. '창의적으로 구조물을 쌓았고'처럼 추상적 표현 대신, "
        "  사용한 재료(컵, 블록, 색종이 등)·순서·모양·동작을 구체적으로 적습니다.\n"
        "  예) '일렬로 컵을 놓은 뒤 사다리꼴 모양으로 컵을 쌓아갔고', '빨간 블록을 밑에 깔고 그 위에 노란 블록을 올렸음'처럼 "
        "  보호자가 장면을 그릴 수 있게 쓰세요.\n"
        "- 아이의 표정·반응·행동, 사용한 교구·재료를 구체적으로 넣어 보호자가 생생히 떠올릴 수 있게 하세요.\n"
        "- 대상은 발달장애 청소년이므로, '성취감을 공유하였음', '협동심을 키웠음', '자신감이 생겼음'처럼 "
        "  추상적·감정 해석 표현은 쓰지 마세요. 대신 관찰 가능한 행동만 적습니다.\n"
        "  예) '성취감을 공유하였음' → '만든 작품을 친구에게 보여주었음', '함께 순서대로 블록을 쌓았음', "
        "'완성 후 웃으며 선생님께 보여주었음'처럼 무엇을 했는지·어떤 모습이었는지 구체적으로 쓰세요.\n"
        "- '이상이 없습니다', '완료되었습니다', '이상 없음' 같은 형식적 표현은 절대 쓰지 마세요.\n"
        "- 사진에서 확인할 수 없는 시간대는 빈 문자열(\"\")로 두세요.\n"
        "- special_note는 실제로 특이한 일이 있을 때만 쓰고, 없으면 반드시 빈 문자열(\"\")로 두세요.\n\n"
        "【필수 JSON 형식 규칙】\n"
        "- 순수 JSON만 출력 (코드 블록·설명 문구 없음)\n"
        "- 각 필드 값은 한 줄(줄바꿈 없는 문자열)로 작성\n"
        "- 값 안에 큰따옴표(\") 사용 금지\n\n"
        '{"activity_1430":"등원 및 자유활동 (시간대 반복 금지, 미확인 시 빈 문자열)",'
        '"activity_1500":"그룹활동 (시간대 반복 금지, 미확인 시 빈 문자열)",'
        '"activity_1600":"교구활동 (시간대 반복 금지, 미확인 시 빈 문자열)",'
        '"activity_1700":"요리활동 및 간식 (시간대 반복 금지, 미확인 시 빈 문자열)",'
        '"activity_1800":"정리 및 하원 (시간대 반복 금지, 미확인 시 빈 문자열)",'
        '"special_note":"실제 특이사항만, 없으면 빈 문자열"}'
    )

    messages = [{
        "role": "user",
        "content": img_content + [{"type": "text", "text": prompt_text}]
    }]

    try:
        raw = _call_openai(api_key, messages, max_tokens=2200)
        data = _extract_json(raw)
        if not data:
            raise ValueError("OpenAI 응답에서 JSON을 파싱할 수 없습니다.")

        _empty = {"없음", "해당없음", "없음.", "해당 없음", "-", "N/A", "n/a",
                  "특이사항 없음", "이상 없음", "특이사항없음", "이상없음"}
        def _skip(v):
            t = v.strip().strip("[]().")
            return not t or t in _empty

        slots = [
            ("activity_1430", "14:30~15:00 등원 및 자유활동"),
            ("activity_1500", "15:00~16:00 그룹활동"),
            ("activity_1600", "16:00~17:00 교구활동"),
            ("activity_1700", "17:00~18:00 요리활동 및 간식"),
            ("activity_1800", "18:00~18:30 정리 및 하원"),
        ]
        full_lines = []
        for key, label in slots:
            val = data.get(key, "").strip()
            # 시간대 접두어 제거
            import re as _re
            val = _re.sub(r'^\d{1,2}:\d{2}[~\-]\d{1,2}:\d{2}\s*', '', val)
            if not _skip(val):
                # 공문서 문체가 섞이면 따뜻한 교사 말투로 후처리
                val = _rewrite_official_style_to_teacher(val)
                # 사용자 요청: '~어요' 마지막 어미를 '~음'으로 변경
                val = _rewrite_teacher_ending_to_eum(val)
                full_lines.append(f"[{label}]\n{val}")
        special = data.get("special_note", "").strip()
        if special and not _skip(special):
            special = _rewrite_official_style_to_teacher(special)
            special = _rewrite_teacher_ending_to_eum(special)
            full_lines.append(f"[특이사항]\n{special}")

        data["full_text"] = "\n\n".join(full_lines)
        return data
    except Exception as e:
        print(f"[openai_service] generate_activity_content 오류: {e}")
        return _demo_activity(photo_metas, error=str(e))


def _demo_activity(photo_metas=None, error=None):
    demo_prefix = "[데모]" if not error else f"[오류: {error[:60]}]"
    metas = photo_metas or []
    places = [m['place'] for m in metas if m.get('place')]
    p = places[0] if places else "올리브청소년방과후센터"
    special_note = " / ".join(m['note'] for m in metas if m.get('note'))
    return {
        "activity_1430": (
            f"{demo_prefix} 아동들이 {p}에 귀원하여 손을 씻고 자유 놀이에 참여했음. "
            "각자 선호하는 활동을 선택하였고, 교사는 개별 아동을 관찰하며 필요 시 언어적 촉진을 도와주었음."
        ),
        "activity_1500": (
            "교사의 안내에 따라 테이블 게임 그룹 활동이 진행됐음. "
            "아동들은 순서를 기다리며 또래와 눈맞춤·언어 교환 등 긍정적인 상호작용을 보였다."
        ),
        "activity_1600": (
            "색종이 접기 및 끼우기 교구를 활용한 소근육 발달 활동을 했음. "
            "아동들은 교사의 시범을 보고 순서에 따라 교구를 조작하며 과제를 완성했음."
        ),
        "activity_1700": (
            "샌드위치 만들기 요리 활동이 진행됐음. 아동들은 식빵·햄·채소 등 재료를 직접 선택해 "
            "적극적으로 참여했으며, 완성된 간식을 함께 나누어 먹었어요."
        ),
        "activity_1800": (
            "활동을 마무리하고 아동들이 개인 소지품을 스스로 정리했음. "
            "교사의 안내에 따라 하원 순서를 기다리며 보호자에게 전해 드렸음."
        ),
        "special_note": special_note if special_note else "특이사항 없음",
        "full_text": (
            f"{demo_prefix} {p}에서 방과후 프로그램이 진행됐음. "
            "학생들은 선생님의 안내에 따라 활동에 적극적으로 참여했으며 즐거운 시간을 보냈음.\n"
            "(실제 운영 시 OpenAI API 키를 설정하면 사진 분석 결과가 자동으로 입력됩니다.)"
        ),
        "_demo": True
    }


# ──────────────────────────────────────────────
# 품위서 – 영수증 OCR + 필드 추출
# ──────────────────────────────────────────────
_RECEIPT_SINGLE_PROMPT = (
    "이 영수증 사진에 **실제로 보이는 내용만** 그대로 추출하세요. "
    "사진에 없는 품목, 금액, 날짜를 만들거나 추측하지 마세요. "
    "상품명·단가·수량·금액은 영수증에 적힌 순서대로 모두 넣고, "
    "날짜는 '구매' 또는 거래일시에 나온 날짜(YYYY-MM-DD), "
    "합계는 '합계' 또는 '결제대상금액'·'소계 금액'에 적힌 **최종 결제 합계** 숫자만 사용하세요.\n\n"
    "다음 JSON 형식으로만 응답하세요. JSON 외 텍스트는 넣지 마세요.\n"
    "{\n"
    '  "purchase_date": "YYYY-MM-DD",\n'
    '  "store_name": "",\n'
    '  "items": [\n'
    '    {"name": "품명", "qty": "수량", "unit": "개", "unit_price": "단가", "amount": "금액", "note": ""}\n'
    "  ],\n"
    '  "total_amount": "합계금액숫자만"\n'
    "}\n"
    '숫자는 쉼표 없이 숫자만(예: 101360), 날짜는 YYYY-MM-DD. unit은 영수증에 없으면 "개".'
)


def _parse_receipt_total_amount(value) -> int:
    """영수증 하단 합계 문자열을 정수로 변환합니다. 파싱 실패 시 0."""
    if value is None:
        return 0
    t = str(value).strip().replace(",", "")
    if not t:
        return 0
    try:
        return int(float(t))
    except ValueError:
        return 0


def _extract_single_receipt_image(image_path: str, api_key: str) -> dict:
    """영수증 이미지 1장만 API에 보내 구조화합니다."""
    img_content = _build_image_content([image_path], detail="high")
    messages = [{
        "role": "user",
        "content": img_content + [{"type": "text", "text": _RECEIPT_SINGLE_PROMPT}],
    }]
    raw = _call_openai(api_key, messages, max_tokens=2000)
    data = _extract_json(raw)
    if not data:
        raise ValueError("JSON 파싱 실패")
    return data


def _merge_multi_receipt_extractions(parts: list) -> dict:
    """
    영수증별 추출 결과를 업로드 순서대로 이어 붙입니다.
    total_amount: 각 장의 하단 합계(total_amount) 숫자를 더한 값.
    purchase_date: 추출된 날짜 중 더 늦은 날(YYYY-MM-DD 문자열 비교).
    store_name: 매장명을 업로드 순서대로 ' | '로 연결(같은 이름이어도 장마다 유지).
    """
    all_items = []
    sum_footer = 0
    dates = []
    stores_ordered = []

    for p in parts:
        for it in p.get("items") or []:
            all_items.append(it)
        sum_footer += _parse_receipt_total_amount(p.get("total_amount"))
        d = (p.get("purchase_date") or "").strip()
        if d:
            dates.append(d)
        s = (p.get("store_name") or "").strip()
        if s:
            stores_ordered.append(s)

    purchase_date = max(dates) if dates else ""
    store_name = " | ".join(stores_ordered)
    total_amount = str(sum_footer) if sum_footer else ""

    return {
        "purchase_date": purchase_date,
        "store_name": store_name,
        "items": all_items,
        "total_amount": total_amount,
    }


def _receipt_part_for_hwp(p: dict) -> dict:
    """HWP 장별 칸(합계·거래처)용 최소 필드."""
    if not p:
        return {"store_name": "", "total_amount": "", "purchase_date": ""}
    return {
        "store_name": (p.get("store_name") or "").strip(),
        "total_amount": (p.get("total_amount") or "").strip(),
        "purchase_date": (p.get("purchase_date") or "").strip(),
    }


def extract_receipt_data(image_paths: list, api_key: str) -> dict:
    """
    영수증 사진에서 품목 정보를 추출합니다.
    - 이미지가 여러 장이면 **장마다 API를 1회씩** 호출한 뒤, 품목 행은 순서대로 합치고
      합계는 각 영수증 하단 합계 숫자의 **합**으로 둡니다.
    Returns dict: {purchase_date, store_name, items, total_amount, receipt_parts}
      receipt_parts: 업로드 순서와 동일한 {store_name, total_amount, purchase_date} 목록
    """
    paths = [p for p in (image_paths or []) if p and os.path.isfile(p)]
    if not paths:
        return _demo_receipt()
    if not api_key:
        return _demo_receipt()

    try:
        if len(paths) == 1:
            one = _extract_single_receipt_image(paths[0], api_key)
            one["receipt_parts"] = [_receipt_part_for_hwp(one)]
            return one
        raw_parts = [_extract_single_receipt_image(p, api_key) for p in paths]
        merged = _merge_multi_receipt_extractions(raw_parts)
        merged["receipt_parts"] = [_receipt_part_for_hwp(p) for p in raw_parts]
        return merged
    except Exception as e:
        print(f"[openai_service] extract_receipt_data 오류: {e}")
        return _demo_receipt()


def _demo_receipt():
    base = {
        "purchase_date": "2026-03-02",
        "store_name": "데모마트",
        "items": [
            {"name": "색연필 세트", "qty": "2", "unit": "개",
             "unit_price": "5000", "amount": "10000", "note": "미술활동용"},
            {"name": "스케치북", "qty": "5", "unit": "권",
             "unit_price": "2000", "amount": "10000", "note": ""},
        ],
        "total_amount": "20000",
        "_demo": True,
    }
    base["receipt_parts"] = [_receipt_part_for_hwp(base)]
    return base


# ──────────────────────────────────────────────
# 계획서 – 목적·목표·프로그램내용·기대효과 생성
# ──────────────────────────────────────────────
def generate_plan_content(image_paths: list, purchase_summary: str, api_key: str) -> dict:
    """
    사진 + 품위서 요약을 바탕으로 계획서 내용을 생성합니다.
    Returns dict: {purpose, goal, program_content, expected_effect}
    """
    if not api_key:
        return _demo_plan(purchase_summary)

    img_content = _build_image_content(image_paths) if image_paths else []
    purchase_text = f"\n\n구매 물품 내역:\n{purchase_summary}" if purchase_summary else ""
    prompt_text = (
        "발달장애를 가진 청소년들을 위한 방과후 프로그램 계획서를 작성해 주세요."
        f"{purchase_text}\n\n"
        "위 사진과 구매 물품을 참고하여, 아래 JSON 형식으로 작성해 주세요. 글자 수를 반드시 지키세요.\n"
        "{\n"
        '  "purpose": "활동 목적 (80자 내외)",\n'
        '  "goal": "활동 목표 (80자 내외, 번호 매기기)",\n'
        '  "program_content": "프로그램 진행 내용 (160자 이내)",\n'
        '  "expected_effect": "기대 효과 (160자 이내)"\n'
        "}\n"
        "JSON만 응답해 주세요."
    )

    messages = [{
        "role": "user",
        "content": img_content + [{"type": "text", "text": prompt_text}]
    }]

    try:
        raw = _call_openai(api_key, messages, max_tokens=800)
        data = _extract_json(raw)
        if not data:
            raise ValueError("JSON 파싱 실패")
        return data
    except Exception as e:
        print(f"[openai_service] generate_plan_content 오류: {e}")
        return _demo_plan(purchase_summary)


def _demo_plan(purchase_summary):
    return {
        "purpose": "발달장애 청소년의 창의적 표현 능력과 사회성 향상을 위한 방과후 예술 프로그램 운영.",
        "goal": "1. 창의적 표현 능력 향상 2. 소근육·집중력 강화 3. 또래 상호작용 증진 4. 정서적 안정",
        "program_content": (
            "1단계: 도입–주제 소개·준비물 확인(10분) "
            "2단계: 본 활동–미술/공예 진행(40분) 3단계: 마무리–감상·정리(10분)"
        ),
        "expected_effect": (
            "참여를 통해 자기표현 능력이 향상되고 성취감·자존감이 증진될 것으로 기대됩니다."
        ),
        "_demo": True
    }
