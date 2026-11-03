import json
import os
import re
import subprocess
from pptx import Presentation
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from xml.sax.saxutils import escape

# ==============================================================================
# 1. 기존 테마 색상 순환 구조 및 템플릿 설정 (기존 기능 100% 유지)
# ==============================================================================
COLOR_CYCLE = ["deep_navy", "burgundy", "forest_green"]

TEMPLATE_MAP = {
    "deep_navy": "assets/templates/이미지카드- 텍스트 첨부(딥네이비).pptx",
    "burgundy": "assets/templates/이미지카드- 텍스트 첨부(버건디).pptx",
    "forest_green": "assets/templates/이미지카드- 텍스트 첨부(포레스트그린).pptx",
}

# ==============================================================================
# 2. 줄바꿈 글자수 제한 상수 (필요 시 숫자만 조절하세요)
# ==============================================================================
SLIDE1_MAIN_TITLE_MAX_CHARS = 7   # 1번 표지 메인 제목 한 줄 최대 글자수
SLIDE1_SUBTITLE_MAX_CHARS = 12    # 1번 표지 부제 한 줄 최대 글자수
SLIDE1_CAPTION_MAX_CHARS = 18     # 1번 표지 하단 요약 한 줄 최대 글자수
BULLET_TOTAL_MAX_CHARS = 28       # 2~5번 슬라이드 불릿 전체(콜론 앞+뒤 합산) 한 줄 최대 글자수


def wrap_by_words(text: str, max_chars: int) -> list[str]:
    """어절(띄어쓰기) 단위로 줄바꿈하며, 한 줄이 max_chars를 넘지 않도록 분할합니다."""
    normalized = "\n".join(re.sub(r"[ \t]+", " ", p.strip()) for p in text.split("\n"))
    lines = []
    for para in normalized.split("\n"):
        if not para:
            continue
        words = para.split(" ")
        current = words[0]
        for w in words[1:]:
            candidate = f"{current} {w}"
            if len(candidate) <= max_chars:
                current = candidate
            else:
                lines.append(current)
                current = w
        if current:
            lines.append(current)
    return lines


def format_slide1_subtitle_specs(
    sub_title: str,
    font_name: str = "G마켓 산스 TTF Bold",
    sz: int = 4400,
    max_chars: int = SLIDE1_SUBTITLE_MAX_CHARS,
) -> list[dict]:
    """
    [수정 1] 1번 슬라이드 부제 라인 정렬:
    앞의 불필요한 콜론(': ')이나 어색한 공백을 정리하고,
    쉼표(,) 및 글자수(max_chars) 기준으로 줄을 나눈 뒤 우측 정렬(align="r") 스펙으로 변환합니다.
    """
    cleaned = re.sub(r"^:\s*", "", sub_title.strip())
    cleaned = re.sub(r"\s+,", ",", cleaned)
    cleaned = re.sub(r"\s+\?", "?", cleaned)

    if "\n" not in cleaned and "," in cleaned and len(cleaned) > max_chars:
        first, rest = cleaned.split(",", 1)
        cleaned = f"{first.strip()},\n{rest.strip()}"

    lines = wrap_by_words(cleaned, max_chars=max_chars)
    return [{"text": line, "font_name": font_name, "align": "r", "sz": sz} for line in lines]


def make_para_xml(
    text: str = "",
    font_name: str = "G마켓 산스 TTF Medium",
    align: str = "l",
    sz: int | None = None,
    bold: bool = False,
    color_hex: str | None = None,
    hidden_prefix: str = "",
) -> str:
    """
    [수정 2] hidden_prefix가 주어지면 앞부분(예: '• 이득금지의 원칙: ')을 투명(<a:noFill/>)으로 먼저 그려서,
    둘째 줄로 내려간 글자('공제')가 첫째 줄 콜론(':') 뒤 첫 글자('노후')와 1픽셀 오차 없이 같은 라인에 정렬되게 합니다.
    """
    sz_attr = f' sz="{sz}"' if sz else ""
    b_attr = ' b="1"' if bold else ""
    fill_xml = (
        f'<a:solidFill><a:srgbClr val="{color_hex}"/></a:solidFill>'
        if color_hex
        else '<a:solidFill><a:schemeClr val="bg1"/></a:solidFill>'
    )
    font_xml = f"""
        <a:latin typeface="{font_name}" panose="02000000000000000000" pitchFamily="2" charset="-127"/>
        <a:ea typeface="{font_name}" panose="02000000000000000000" pitchFamily="2" charset="-127"/>
    """
    if text == "" and hidden_prefix == "":
        return (
            f'<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            f'<a:pPr algn="{align}"/>'
            f'<a:endParaRPr lang="ko-KR" altLang="en-US"{sz_attr}{b_attr} dirty="0">{fill_xml}{font_xml}</a:endParaRPr>'
            f"</a:p>"
        )

    hidden_run_xml = ""
    if hidden_prefix:
        hidden_run_xml = (
            f"<a:r>"
            f'<a:rPr lang="ko-KR" altLang="en-US"{sz_attr}{b_attr} dirty="0"><a:noFill/>{font_xml}</a:rPr>'
            f"<a:t>{escape(hidden_prefix)}</a:t>"
            f"</a:r>"
        )

    visible_run_xml = (
        f"<a:r>"
        f'<a:rPr lang="ko-KR" altLang="en-US"{sz_attr}{b_attr} dirty="0">{fill_xml}{font_xml}</a:rPr>'
        f"<a:t>{escape(text)}</a:t>"
        f"</a:r>"
    )

    return (
        f'<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<a:pPr algn="{align}"/>'
        f"{hidden_run_xml}{visible_run_xml}"
        f'<a:endParaRPr lang="ko-KR" altLang="en-US"{sz_attr}{b_attr} dirty="0">{fill_xml}{font_xml}</a:endParaRPr>'
        f"</a:p>"
    )


def expand_bullet_to_specs(
    text: str,
    font_name: str = "G마켓 산스 TTF Medium",
    sz: int | None = None,
    bold: bool = False,
    color_hex: str | None = None,
    total_max_chars: int = BULLET_TOTAL_MAX_CHARS,
) -> list[dict]:
    """
    문장이 일정 글자수(total_max_chars)를 넘어 여러 줄로 나뉠 때:
    1) 문장에 콜론(':')이 있으면 콜론 앞부분('• 이득금지의 원칙: ')을 기준으로 계산하여,
       둘째 줄부터는 hidden_prefix에 '• 이득금지의 원칙: '을 넣어 콜론 뒤 첫 글자 라인에 완벽 정렬합니다.
    2) 콜론이 없고 기호('• ', '1. ', '✔ ', '💡 ')만 있으면 기호 뒤 첫 글자 라인에 정렬합니다.
    """
    clean_text = re.sub(r"[ \t]+", " ", text.strip())

    # 1. 콜론(':')이 포함된 문장 처리
    if ":" in clean_text:
        prefix_part, body_part = clean_text.split(":", 1)
        prefix = f"{prefix_part.strip()}: "
        body = body_part.strip()

        # 콜론 앞부분이 너무 길지 않을 때 콜론 뒤 내어쓰기 적용
        body_max_chars = max(10, total_max_chars - len(prefix))
        body_lines = wrap_by_words(body, max_chars=body_max_chars)

        specs = []
        for idx, b_line in enumerate(body_lines):
            if idx == 0:
                specs.append(
                    {
                        "text": f"{prefix}{b_line}",
                        "font_name": font_name,
                        "sz": sz,
                        "bold": bold,
                        "color_hex": color_hex,
                    }
                )
            else:
                specs.append(
                    {
                        "text": b_line,
                        "hidden_prefix": prefix,
                        "font_name": font_name,
                        "sz": sz,
                        "bold": bold,
                        "color_hex": color_hex,
                    }
                )
        return specs

    # 2. 콜론이 없는 불릿/기호 문장 처리 ('• ', '1. ', '✔ ', '💡 ')
    match = re.match(r"^([•✔💡]|\d+\.)\s+", clean_text)
    if match:
        marker = match.group(0)
        body = clean_text[len(marker) :].strip()
        body_max_chars = max(12, total_max_chars - len(marker))
        body_lines = wrap_by_words(body, max_chars=body_max_chars)

        specs = []
        for idx, b_line in enumerate(body_lines):
            if idx == 0:
                specs.append(
                    {
                        "text": f"{marker}{b_line}",
                        "font_name": font_name,
                        "sz": sz,
                        "bold": bold,
                        "color_hex": color_hex,
                    }
                )
            else:
                specs.append(
                    {
                        "text": b_line,
                        "hidden_prefix": marker,
                        "font_name": font_name,
                        "sz": sz,
                        "bold": bold,
                        "color_hex": color_hex,
                    }
                )
        return specs

    # 3. 일반 문장
    lines = wrap_by_words(clean_text, max_chars=total_max_chars)
    return [
        {"text": line, "font_name": font_name, "sz": sz, "bold": bold, "color_hex": color_hex}
        for line in lines
    ]


def replace_paragraphs(shape, specs):
    txBody = shape.text_frame._txBody
    for p in list(txBody.findall(qn("a:p"))):
        txBody.remove(p)
    for spec in specs:
        txBody.append(parse_xml(make_para_xml(**spec)))


def get_next_theme(current_theme):
    if current_theme in COLOR_CYCLE:
        idx = COLOR_CYCLE.index(current_theme)
        return COLOR_CYCLE[(idx + 1) % len(COLOR_CYCLE)]
    return COLOR_CYCLE[0]


def main():
    json_path = "posts/current_post.json"
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # 1. 현재 지정된 테마 색상 확인 (없으면 기본 deep_navy, 'navy'로 적혀 있어도 'deep_navy'로 호환)
    current_theme = data.get("theme_color", "deep_navy")
    if current_theme == "navy":
        current_theme = "deep_navy"

    # 2. 이번에 사용할 테마는 현재 지정된 색상 그대로 사용
    tpl_path = TEMPLATE_MAP.get(current_theme, TEMPLATE_MAP["deep_navy"])
    print(f"🎨 선택된 카드뉴스 테마 색상: {current_theme}")

    prs = Presentation(tpl_path)
    cards = data["cards"]

    # 1번 슬라이드 (표지: 메인 타이틀 & 부제 글자수 제한 우측 정렬 + 하단 캡션 줄바꿈 정렬)
    s1 = prs.slides[0]
    main_title_lines = wrap_by_words(cards["slide1"]["main_title"], max_chars=SLIDE1_MAIN_TITLE_MAX_CHARS)
    replace_paragraphs(
        s1.shapes[0],
        [{"text": line, "font_name": "G마켓 산스 TTF Bold", "align": "r", "sz": 6000} for line in main_title_lines],
    )
    replace_paragraphs(
        s1.shapes[2],
        format_slide1_subtitle_specs(cards["slide1"]["sub_title"], font_name="G마켓 산스 TTF Bold", sz=4400),
    )
    caption_specs = []
    for line in cards["slide1"]["left_caption"]:
        for wrapped_line in wrap_by_words(line, max_chars=SLIDE1_CAPTION_MAX_CHARS):
            caption_specs.append({"text": wrapped_line, "font_name": "G마켓 산스 TTF Medium", "align": "l"})
    replace_paragraphs(s1.shapes[1], caption_specs)

    # 2~4번 슬라이드 (본문 요약 1~3: 콜론 ':' 기준 내어쓰기 정렬 적용)
    for idx, item in enumerate(cards["slides_2_to_4"], start=1):
        slide = prs.slides[idx]
        title_specs = expand_bullet_to_specs(
            item["title"], font_name="G마켓 산스 TTF Bold", sz=4500, total_max_chars=18
        )
        replace_paragraphs(slide.shapes[0], title_specs)

        body_specs = [{"text": item["headline"]}, {"text": ""}]
        for b_line in item["bullets"]:
            body_specs.extend(expand_bullet_to_specs(b_line, total_max_chars=BULLET_TOTAL_MAX_CHARS))
        body_specs.append({"text": ""})
        body_specs.extend(
            expand_bullet_to_specs(
                item["note"], bold=True, color_hex="92D050", total_max_chars=BULLET_TOTAL_MAX_CHARS
            )
        )
        replace_paragraphs(slide.shapes[1], body_specs)

    # 5번 슬라이드 (핵심 정리 및 소비자선임권 고정 안내: 콜론 ':' 및 기호 기준 내어쓰기 적용)
    s5 = prs.slides[4]
    s5_title_specs = expand_bullet_to_specs(
        cards["slide5"]["title"], font_name="G마켓 산스 TTF Bold", sz=4800, total_max_chars=18
    )
    replace_paragraphs(s5.shapes[0], s5_title_specs)

    s5_specs = [{"text": cards["slide5"]["headline"]}, {"text": ""}]
    for b_line in cards["slide5"]["bullets"]:
        s5_specs.extend(expand_bullet_to_specs(b_line, total_max_chars=BULLET_TOTAL_MAX_CHARS))
    s5_specs.append({"text": ""})
    for g_line in cards["slide5"]["green_notes"]:
        s5_specs.extend(
            expand_bullet_to_specs(
                g_line, bold=True, color_hex="92D050", total_max_chars=BULLET_TOTAL_MAX_CHARS
            )
        )
    replace_paragraphs(s5.shapes[1], s5_specs)

    os.makedirs("posts/images", exist_ok=True)
    out_pptx = "posts/images/generated_cards.pptx"
    prs.save(out_pptx)

    # PPTX -> PDF -> 고화질 PNG(1080x1080) 5장 자동 변환
    subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "pdf", out_pptx, "--outdir", "posts/images"],
        check=True,
    )
    subprocess.run(
        ["pdftoppm", "-png", "-r", "96", "posts/images/generated_cards.pdf", "posts/images/slide"],
        check=True,
    )

    # 3. 다음 실행을 위해 current_post.json의 theme_color를 자동으로 다음 순서 색상으로 업데이트 후 푸시
    next_theme = get_next_theme(current_theme)
    data["theme_color"] = next_theme
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"🔄 다음 글을 위해 색상을 자동으로 [{next_theme}](으)로 갱신했습니다.")


if __name__ == "__main__":
    main()
