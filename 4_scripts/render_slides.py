import os
import asyncio
import json
from jinja2 import Environment, FileSystemLoader
from playwright.async_api import async_playwright

async def render_slides():
    print("--- Start Rendering HTML Slides to PNG ---")
    
    # 1. 데이터 로드
    hospital_data_path = "0_input/hospital_data.json"
    with open(hospital_data_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    hospital_id = data.get("hospital_id", "unknown")
    hospital_name = data.get("hospital_name", "병원")
    
    # 출력 경로 설정
    output_dir = f"1_assets/hospitals/{hospital_id}/renders"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    # 2. Jinja2 설정
    template_dir = "3_templates/html_layouts"
    env = Environment(loader=FileSystemLoader(template_dir))
    
    # 3. Playwright 시작
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page(viewport={"width": 1280, "height": 720})
        
        # 렌더링할 슬라이드 목록 정의
        slides_to_render = [
            {"num": 1, "template": "slide_01_cover.html", "ctx": data},
            {"num": 2, "template": "slide_02_summary.html", "ctx": data},
            {"num": 3, "template": "slide_03_overview.html", "ctx": data},
            {"num": 4, "template": "slide_04_test.html", "ctx": {**data, "test": data["ai_tests"][0], "test_index": 1, "screenshot_path": os.path.abspath(f"1_assets/hospitals/{hospital_id}/04_재미나이아현비뇨기과추천.png")}},
            {"num": 5, "template": "slide_04_test.html", "ctx": {**data, "test": data["ai_tests"][1], "test_index": 2}},
            {"num": 6, "template": "slide_04_test.html", "ctx": {**data, "test": data["ai_tests"][2], "test_index": 3, "screenshot_path": os.path.abspath(f"1_assets/hospitals/{hospital_id}/06_퍼플렉시티요로결석쇄석관련질문.png")}},
            {"num": 7, "template": "slide_07_logic.html", "ctx": data},
            {"num": 8, "template": "slide_08_competitor.html", "ctx": data},
            {"num": 9, "template": "slide_09_diagnosis.html", "ctx": data},
            {"num": 10, "template": "slide_10_issues.html", "ctx": data},
            {"num": 11, "template": "slide_11.html", "ctx": data},
            {"num": 12, "template": "slide_12.html", "ctx": data},
            {"num": 13, "template": "slide_13.html", "ctx": data},
            {"num": 14, "template": "slide_14.html", "ctx": data},
            {"num": 15, "template": "slide_15_strategy.html", "ctx": data},
            {"num": 16, "template": "slide_16_core.html", "ctx": data},
            {"num": 17, "template": "slide_17_faq.html", "ctx": data},
            {"num": 18, "template": "slide_18_tech.html", "ctx": data},
            {"num": 19, "template": "slide_19_process.html", "ctx": data},
            {"num": 20, "template": "slide_20_closing.html", "ctx": data},
        ]
        
        for slide in slides_to_render:
            num = slide["num"]
            template_name = slide["template"]
            # 딕셔너리 복사하여 슬라이드별 고유 번호 부여
            ctx = slide["ctx"].copy() if isinstance(slide["ctx"], dict) else {}
            ctx["slide_number"] = num
            
            print(f"Rendering Slide {num}...")
            
            # HTML 렌더링
            template = env.get_template(template_name)
            html_content = template.render(**ctx)
            
            # 절대 경로로 이미지 참조를 위해 HTML 수정 (또는 파일로 저장 후 로드)
            # 여기서는 파일로 임시 저장하고 로컬 파일 URL로 엽니다.
            temp_html = f"tmp_slide_{num}.html"
            with open(os.path.join(template_dir, temp_html), 'w', encoding='utf-8') as f:
                f.write(html_content)
                
            file_url = f"file:///{os.path.abspath(os.path.join(template_dir, temp_html))}"
            await page.goto(file_url)
            
            # 배경만 캡처 (텍스트 숨김)
            await page.add_script_tag(content="document.body.classList.add('bg-only');")
            bg_output_path = os.path.join(output_dir, f"bg_{num:02d}.png")
            await page.screenshot(path=bg_output_path, full_page=False)
            
            # 원래 상태로 복구 (또는 그냥 다음 슬라이드)
            await page.add_script_tag(content="document.body.classList.remove('bg-only');")
            
            # 임시 파일 삭제
            os.remove(os.path.join(template_dir, temp_html))
            
        await browser.close()
        
    print(f"\n--- Success! Rendered PNGs in: {output_dir} ---")

if __name__ == "__main__":
    asyncio.run(render_slides())
