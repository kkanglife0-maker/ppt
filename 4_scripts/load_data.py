import os
from utils import load_json

def load_hospital_data(file_path="0_input/hospital_data.json"):
    """병원 데이터 JSON을 로드합니다."""
    data = load_json(file_path)
    if not data:
        raise FileNotFoundError(f"필수 파일 {file_path}를 찾을 수 없습니다.")
    
    # 필수 필드 체크
    required_fields = ["hospital_id", "hospital_name", "ai_tests", "ai_analysis_summary"]
    missing = [f for f in required_fields if f not in data]
    if missing:
        print(f"Warning: 필수 데이터 누락 - {', '.join(missing)}")
        
    return data

def load_style_config(file_path="3_templates/style_config.json"):
    """스타일 설정 JSON을 로드합니다."""
    return load_json(file_path)

def load_fixed_slides(directory="1_assets/fixed_slides"):
    """고정 슬라이드 JSON 파일들을 로드합니다."""
    fixed_slides = {}
    if not os.path.exists(directory):
        print(f"Warning: 고정 슬라이드 디렉토리 누락 - {directory}")
        return fixed_slides
        
    for filename in os.listdir(directory):
        if filename.endswith(".json"):
            slide_data = load_json(os.path.join(directory, filename))
            if slide_data and "slide_number" in slide_data:
                fixed_slides[str(slide_data["slide_number"])] = slide_data
                
    return fixed_slides
