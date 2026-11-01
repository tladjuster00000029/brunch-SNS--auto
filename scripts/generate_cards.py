import json
import os
import subprocess
from pptx import Presentation
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from xml.sax.saxutils import escape

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

def main():
    with open("posts/current_post.json", "r", encoding="utf-8") as f:
        data = json.load(f)

    theme = data.get("theme_color", "deep_navy")
    tpl_path = TEMPLATE_MAP.get(theme, TEMPLATE_MAP["deep_navy"])
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
    print("✅ 카드뉴스 PNG 5장 및 PPTX 자동 생성이 완료되었습니다!")

if __name__ == "__main__":
    main()
