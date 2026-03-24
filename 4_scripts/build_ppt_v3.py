import os
import json
from pptx import Presentation
from pptx.util import Inches
from slide_factory import SlideFactory

def build_ppt_v3():
    print("--- Start Building Editable Premium PPT (v3) ---")
    
    # 1. 데이터 로드
    with open("0_input/hospital_data.json", "r", encoding="utf-8") as f:
        hospital_data = json.load(f)
    
    with open("3_templates/style_config.json", "r", encoding="utf-8") as f:
        style_config = json.load(f)
        
    hospital_name = hospital_data.get("hospital_name", "병원")
    output_dir = "5_output"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    ppt_filename = f"{hospital_name}_AEO_Report_Premium_Editable.pptx"
    ppt_path = os.path.join(output_dir, ppt_filename)
    
    # 2. PPT 초기화
    prs = Presentation()
    # 16:9 설정
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    factory = SlideFactory(prs, style_config, hospital_data)
    
    # 3. 슬라이드 생성 루프 (1~20)
    layout_dir = "3_templates/ppt_layouts"
    for i in range(1, 21):
        layout_path = os.path.join(layout_dir, f"slide_{i:02d}.json")
        
        if os.path.exists(layout_path):
            print(f"Generating Slide {i} from JSON layout...")
            with open(layout_path, 'r', encoding='utf-8') as f:
                slide_layout = json.load(f)
            factory.add_object_slide(slide_layout)
        else:
            print(f"Slide {i} layout not found, using default dynamic slide...")
            factory.add_dynamic_slide(i)

    # 4. 저장
    prs.save(ppt_path)
    print(f"\n--- Success! Editable Premium PPT Created: {ppt_path} ---")

if __name__ == "__main__":
    build_ppt_v3()
