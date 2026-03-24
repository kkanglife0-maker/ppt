import os
import json
from datetime import datetime

def load_json(file_path):
    """지정된 경로의 JSON 파일을 읽어서 반환합니다."""
    if not os.path.exists(file_path):
        print(f"Warning: File not found at {file_path}")
        return None
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Error loading JSON from {file_path}: {e}")
        return None

def get_output_filename(hospital_id):
    """hospital_id_YYYYMMDD.pptx 형식의 파일명을 생성합니다."""
    date_str = datetime.now().strftime("%Y%m%d")
    return f"{hospital_id}_{date_str}.pptx"

def ensure_dir(directory):
    """디렉토리가 없으면 생성합니다."""
    if not os.path.exists(directory):
        os.makedirs(directory)

def get_image_path(hospital_id, image_filename):
    """병원별 이미지 폴더에서 이미지의 상대 경로를 반환합니다."""
    base_path = os.path.join("1_assets", "hospitals", hospital_id)
    return os.path.join(base_path, image_filename)
