import json
import os
import subprocess
from pptx import Presentation
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from xml.sax.saxutils import escape

# 사용 가능한 색상 순서 정의 (순환 구조)
COLOR_CYCLE = ["deep_navy", "burgundy", "forest_green"]

TEMPLATE_MAP = {
    "deep_navy": "assets/templates/이미지카드- 텍스트 첨부(딥네이비).pptx",
    "burgundy": "assets/templates/이미지카드- 텍스트 첨부(버건디).pptx",
    "forest_green": "assets/templates/이미지카드- 텍스트 첨부(포레스트그린).pptx",
}

def make_para_xml(text, font_name="G마켓 산스 TTF Medium", align="l", sz=None, bold=False, color_hex=None):
    sz_attr = f' sz="{sz}"' if sz else ''
    b_attr = ' b="1"' if bold else ''
    fill_xml = f'<a:solidFill><a:srgbClr val="{color_hex}"/></a:solidFill>' if color_hex else '<a:solidFill><a:schemeClr val="bg1"/></a:solidFill>'
    font_xml = f'''
        <a:latin typeface="{font_name}" panose="02000000000000000000" pitchFamily="2" charset="-127"/>
        <a:ea typeface="{font_name}" panose="02000000000000000000" pitchFamily="2" charset="-127"/>
    '''
    if text == "":
        return f'<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:pPr algn="{align}"/><a:endParaRPr lang="ko-KR" altLang="en-US"{sz_attr}{b_attr} dirty="0">{fill_xml}{font_xml}</a:endParaRPr></a:p>'
    return f'<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:pPr algn="{align}"/><a:r><a:rPr lang="ko-KR" altLang="en-US"{sz_attr}{b_attr} dirty="0">{fill_xml}{font_xml}</a:rPr><a:t>{escape(text)}</a:t></a:r><a:endParaRPr lang="ko-KR" altLang="en-US"{sz_attr}{b_attr} dirty="0">{fill_xml}{font_xml}</a:endParaRPr></a:p>'

def replace_paragraphs(shape, specs):
    txBody = shape.text_frame._txBody
    for p in list(txBody.findall(qn('a:p'))):
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

    # 1. 현재 지정된 테마 색상 확인 (없으면 기본 deep_navy)
    current_theme = data.get("theme_color", "deep_navy")
    
    # 2. 이번에 사용할 테마는 현재 지정된 색상 그대로 사용
    tpl_path = TEMPLATE_MAP.get(current_theme, TEMPLATE_MAP["deep_navy"])
    print(f"🎨 선택된 카드뉴스 테마 색상: {current_theme}")

    prs = Presentation(tpl_path)
    cards = data["cards"]

    # 1번 슬라이드 (표지)
    s1 = prs.slides[0]
    replace_paragraphs(s1.shapes[0], [{"text": cards["slide1"]["main_title"], "font_name": "G마켓 산스 TTF Bold", "align": "r", "sz": 6000}])
    replace_paragraphs(s1.shapes[2], [{"text": cards["slide1"]["sub_title"], "font_name": "G마켓 산스 TTF Bold", "align": "r", "sz": 4400}])
    replace_paragraphs(s1.shapes[1], [{"text": line, "font_name": "G마켓 산스 TTF Medium", "align": "l"} for line in cards["slide1"]["left_caption"]])

    # 2~4번 슬라이드 (본문 요약 1~3)
    for idx, item in enumerate(cards["slides_2_to_4"], start=1):
        slide = prs.slides[idx]
        replace_paragraphs(slide.shapes[0], [{"text": item["title"], "font_name": "G마켓 산스 TTF Bold", "align": "l", "sz": 4500}])
        body_specs = [{"text": item["headline"]}, {"text": ""}]
        for b_line in item["bullets"]:
            body_specs.append({"text": b_line})
        body_specs.append({"text": ""})
        body_specs.append({"text": item["note"], "bold": True, "color_hex": "92D050"})
        replace_paragraphs(slide.shapes[1], body_specs)

    # 5번 슬라이드 (핵심 정리 및 소비자선임권 고정 안내)
    s5 = prs.slides[4]
    replace_paragraphs(s5.shapes[0], [{"text": cards["slide5"]["title"], "font_name": "G마켓 산스 TTF Bold", "align": "l", "sz": 4800}])
    s5_specs = [{"text": cards["slide5"]["headline"]}, {"text": ""}]
    for b_line in cards["slide5"]["bullets"]:
        s5_specs.append({"text": b_line})
    s5_specs.append({"text": ""})
    for g_line in cards["slide5"]["green_notes"]:
        s5_specs.append({"text": g_line, "bold": True, "color_hex": "92D050"})
    replace_paragraphs(s5.shapes[1], s5_specs)

    os.makedirs("posts/images", exist_ok=True)
    out_pptx = "posts/images/generated_cards.pptx"
    prs.save(out_pptx)

    # PPTX -> PDF -> 고화질 PNG(1080x1080) 5장 자동 변환
    subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", out_pptx, "--outdir", "posts/images"], check=True)
    subprocess.run(["pdftoppm", "-png", "-r", "96", "posts/images/generated_cards.pdf", "posts/images/slide"], check=True)

    # 3. 다음 실행을 위해 current_post.json의 theme_color를 자동으로 다음 순서 색상으로 업데이트 후 푸시
    next_theme = get_next_theme(current_theme)
    data["theme_color"] = next_theme
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"🔄 다음 글을 위해 색상을 자동으로 [{next_theme}](으)로 갱신했습니다.")

if __name__ == "__main__":
    main()
