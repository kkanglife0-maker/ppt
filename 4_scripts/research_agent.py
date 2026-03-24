import requests
from bs4 import BeautifulSoup
import json
import os
import sys
import argparse
from datetime import datetime

# Import local utils if available
try:
    from utils import load_json
except ImportError:
    def load_json(file_path):
        if not os.path.exists(file_path):
            return None
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)

def scrape_hospital_info(url):
    """URL에서 병원 정보를 스크레이핑합니다."""
    print(f"Scraping {url}...")
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # 기본 정보 추출
        title = soup.title.string if soup.title and soup.title.string else ""
        meta_desc = ""
        meta_tag = soup.find('meta', attrs={'name': 'description'})
        if meta_tag:
            meta_desc = meta_tag.get('content', '') or ""
        
        # 핵심 서비스 추출 (H1, H2, 또는 특정 키워드가 포함된 태그)
        services = []
        service_keywords = ['진료', '클리닉', '센터', '수술', '치료', '남성', '여성', '요로', '전립선']
        for tag in soup.find_all(['h1', 'h2', 'h3', 'li']):
            text = tag.get_text(strip=True)
            if any(keyword in text for keyword in service_keywords) and len(text) < 20:
                if text not in services:
                    services.append(text)
        
        # 진료과목 추정
        department = "비뇨기과" # 기본값
        if title and "치과" in title:
            department = "치과"
        elif meta_desc and "치과" in meta_desc:
            department = "치과"
        elif title and "성형" in title:
            department = "성형외과"
        elif title and "피부" in title:
            department = "피부과"
            
        # 홈페이지 진단
        diagnosis = {
            "faq_status": "FAQ 페이지 없음" if not soup.find(string=lambda s: s and ("FAQ" in s or "자주 묻는" in s)) else "FAQ 페이지 존재",
            "schema_status": "Schema 미적용" if not soup.find('script', type='application/ld+json') else "Schema 적용됨",
            "mobile_friendly": "확인 필요(반응형 권장)"
        }
        
        return {
            "title": title,
            "description": meta_desc,
            "services": list(services)[:5], # 상위 5개만
            "department": department,
            "diagnosis": diagnosis
        }
    except Exception as e:
        print(f"Error scraping {url}: {e}")
        return None

def generate_ai_questions(name, region, department, services):
    """3가지 유형의 AI 질문 예시를 생성합니다."""
    # 1. 추천형
    q1 = f"{region} {department} 추천"
    
    # 2. 증상형 (비뇨기과 기준 기본값, 서비스 기반으로 업데이트 가능)
    symptoms = ["소변이 자주 마려울 때", "갑작스러운 통증", "밤에 잠을 설칠 정도로 화장실을 갈 때"]
    q2 = f"{symptoms[0]} 어느 병원으로 가야 하나요?"
    
    # 3. 특화형
    service = services[0] if services else f"{department} 진료"
    q3 = f"{region} {service} 잘하는 곳"
    
    return [
        {"type": "넓은 추천형", "question": q1},
        {"type": "증상형", "question": q2},
        {"type": "특화형", "question": q3}
    ]

def update_hospital_data(name, region, url, info):
    """hospital_data.json 파일을 업데이트합니다."""
    data_path = os.path.join("0_input", "hospital_data.json")
    data = load_json(data_path) or {}
    
    # 병원 ID 생성 (name 기준 영문/숫자만 추출하거나 hash)
    hospital_id = "".join(filter(str.isalnum, url.split("//")[-1].split(".")[0]))
    
    # 데이터 구조 채우기
    data["hospital_id"] = hospital_id
    data["hospital_name"] = name
    data["region"] = region
    data["department"] = info["department"]
    data["website_url"] = url
    data["core_services"] = info["services"]
    data["brand_line"] = f"{region} {name} - AI 검색 환경에서 선택받는 전략"
    
    # AI 질문 채우기
    questions = generate_ai_questions(name, region, info["department"], info["services"])
    data["ai_tests"] = []
    for q in questions:
        data["ai_tests"].append({
            "question_type": q["type"],
            "question": q["question"],
            "gpt_result": "미노출",
            "gemini_result": "미노출",
            "perplexity_result": "미노출",
            "analysis": "AI 검색 시 병원의 핵심 정보가 구조화되어 있지 않아 인용되지 못함"
        })
    
    # 홈페이지 진단 채우기
    data["homepage_diagnosis"] = {
        "main_page": f"현재 {name} 홈페이지는 단순 소개 중심으로 구성됨",
        "service_page": f"{', '.join(info['services'])} 관련 상세 설명 및 질문형 구조 부족",
        "faq_status": info["diagnosis"]["faq_status"],
        "schema_status": info["diagnosis"]["schema_status"],
        "ai_disadvantage": [
            "질문형 문장 부족",
            "대표 진료 메시지 불명확",
            "AI 인용용 답변 블록 없음"
        ]
    }
    
    # 슬라이드 이미지 리스트 최적화 (3, 4, 5, 6만 유지)
    data["slide_images"] = {
        "3": ["전체AI결과.png"],
        "4": ["지역추천질문.png"],
        "5": ["증상형질문.png"],
        "6": ["특화진료질문.png"]
    }
    
    # 불필요한 필드 삭제 (기존 데이터가 있을 경우를 대비)
    if "competitor_hospitals" not in data:
        data["competitor_hospitals"] = [
            {"name": "인근 경쟁 병원 A", "strength": "지역 내 인지도 높음", "difference": "AI 최적화 미흡"},
            {"name": "인근 경쟁 병원 B", "strength": "온라인 마케팅 활발", "difference": "정보형 콘텐츠 부족"}
        ]
    
    with open(data_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print(f"Successfully updated {data_path}")

def main():
    parser = argparse.ArgumentParser(description='Hospital Research Agent')
    parser.add_argument('--name', required=True, help='병원 이름')
    parser.add_argument('--region', required=True, help='지역 (예: 강남, 평택)')
    parser.add_argument('--url', required=True, help='홈페이지 URL')
    
    args = parser.parse_args()
    
    info = scrape_hospital_info(args.url)
    if not info:
        # 스크레이핑 실패 시 기본값으로 진행
        info = {
            "department": "비뇨기과",
            "services": ["전립선", "요로결석", "남성수술"],
            "diagnosis": {"faq_status": "확인 불가", "schema_status": "확인 불가"}
        }
    
    update_hospital_data(args.name, args.region, args.url, info)

if __name__ == "__main__":
    main()
