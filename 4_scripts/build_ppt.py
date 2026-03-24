import os
import sys
from pptx import Presentation
from pptx.util import Inches

# Import local modules
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from utils import get_output_filename, ensure_dir
from load_data import load_hospital_data, load_style_config, load_fixed_slides
from slide_factory import SlideFactory

def main():
    print("--- AEO Hospital PPT Generator Start ---")
    
    try:
        # 1. 데이터 로드
        hospital_data = load_hospital_data()
        style_config = load_style_config()
        fixed_slides = load_fixed_slides()
        
        hospital_id = hospital_data.get("hospital_id", "unknown")
        
        # 2. PPT 초기화 (16:9)
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        
        factory = SlideFactory(prs, style_config, hospital_data)
        
        # 3. 1~20 슬라이드 순서대로 생성
        for i in range(1, 21):
            slide_num_str = str(i)
            print(f"Generating Slide {i}...")
            
            if slide_num_str in fixed_slides:
                # 고정 슬라이드 처리
                slide_data = fixed_slides[slide_num_str]
                if i == 1 and slide_data.get("layout") != "objects":
                    factory.add_title_slide(slide_data)
                elif slide_data.get("layout") == "objects":
                    factory.add_object_slide(slide_data)
                else:
                    factory.add_bullet_slide(slide_data)
            else:
                # 동적 슬라이드 처리
                factory.add_dynamic_slide(i)
        
        # 4. 저장
        output_dir = "5_output"
        ensure_dir(output_dir)
        filename = get_output_filename(hospital_id)
        output_path = os.path.join(output_dir, filename)
        
        prs.save(output_path)
        print(f"\n--- Success! PPT Created: {output_path} ---")
        
    except Exception as e:
        print(f"\n--- Error occurred: {e} ---")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
