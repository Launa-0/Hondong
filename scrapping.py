from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import json
import pandas as pd
import time
import re
import os

# -----------------------------
# 크롬 드라이버 설정
# -----------------------------
driver_path = "C:/Users/USER/Downloads/chromedriver-win64/chromedriver.exe"

service = Service(driver_path)
driver = webdriver.Chrome(service=service)
# -----------------------------
# 스크래핑 대상 URL 리스트
# -----------------------------
urls = {
    "남현": "https://www.daangn.com/kr/realty/?areaViewType=p&in=%EB%82%A8%ED%98%84%EB%8F%99-350&salesType=officetel%2Cone_room%2Ctwo_room&tradeType=borrow",
    "청룡": "https://www.daangn.com/kr/realty/?areaViewType=p&in=%EC%B2%AD%EB%A3%A1%EB%8F%99-346&salesType=officetel%2Cone_room%2Ctwo_room&tradeType=borrow",
    "성현": "https://www.daangn.com/kr/realty/?areaViewType=p&in=%EC%84%B1%ED%98%84%EB%8F%99-343&salesType=officetel%2Cone_room%2Ctwo_room&tradeType=borrow",
    "조원": "https://www.daangn.com/kr/realty/?areaViewType=p&in=%EC%A1%B0%EC%9B%90%EB%8F%99-357&salesType=officetel%2Cone_room%2Ctwo_room&tradeType=borrow",
    "미성": "https://www.daangn.com/kr/realty/?areaViewType=p&in=%EB%AF%B8%EC%84%B1%EB%8F%99-360&salesType=officetel%2Cone_room%2Ctwo_room&tradeType=borrow",
    "서림": "https://www.daangn.com/kr/realty/?areaViewType=p&in=%EC%84%9C%EB%A6%BC%EB%8F%99-353&salesType=officetel%2Cone_room%2Ctwo_room&tradeType=borrow",
    "보라매": "https://www.daangn.com/kr/realty/?areaViewType=p&in=%EB%B3%B4%EB%9D%BC%EB%A7%A4%EB%8F%99-341&salesType=officetel%2Cone_room%2Ctwo_room&tradeType=borrow",
    "신림": "https://www.daangn.com/kr/realty/?areaViewType=p&in=%EC%8B%A0%EB%A6%BC%EB%8F%99-355&salesType=officetel%2Ctwo_room%2Cone_room&tradeType=borrow",
    "봉천": "https://www.daangn.com/kr/realty/?areaViewType=p&in=%EB%B4%89%EC%B2%9C%EB%8F%99-6058&salesType=officetel%2Ctwo_room%2Cone_room&tradeType=borrow",
    "대화": "https://www.daangn.com/kr/realty/?areaViewType=p&in=%EB%8C%80%ED%95%99%EB%8F%99-358&salesType=officetel%2Ctwo_room%2Cone_room&tradeType=borrow",
    "낙성대": "https://www.daangn.com/kr/realty/?areaViewType=p&in=%EB%82%99%EC%84%B1%EB%8C%80%EB%8F%99-345&salesType=one_room%2Ctwo_room%2Cofficetel&tradeType=borrow",
    "난곡": "https://www.daangn.com/kr/realty/?areaViewType=p&in=%EB%82%9C%EA%B3%A1%EB%8F%99-361&salesType=officetel%2Ctwo_room%2Cone_room&tradeType=borrow",
    "은천": "https://www.daangn.com/kr/realty/?areaViewType=p&in=%EC%9D%80%EC%B2%9C%EB%8F%99-347&salesType=officetel%2Ctwo_room%2Cone_room&tradeType=borrow",
    "행운": "https://www.daangn.com/kr/realty/?areaViewType=p&in=%ED%96%89%EC%9A%B4%EB%8F%99-344&salesType=officetel%2Ctwo_room%2Cone_room&tradeType=borrow",
    "인헌": "https://www.daangn.com/kr/realty/?areaViewType=p&in=%EC%9D%B8%ED%97%8C%EB%8F%99-349&salesType=officetel%2Ctwo_room%2Cone_room&tradeType=borrow",
    "신사": "https://www.daangn.com/kr/realty/?areaViewType=p&in=%EC%8B%A0%EC%82%AC%EB%8F%99-354&salesType=one_room%2Ctwo_room%2Cofficetel&tradeType=borrow",
    "서원": "https://www.daangn.com/kr/realty/?areaViewType=p&in=%EC%84%9C%EC%9B%90%EB%8F%99-351&salesType=one_room%2Ctwo_room%2Cofficetel&tradeType=borrow",
    "삼성": "https://www.daangn.com/kr/realty/?areaViewType=p&in=%EC%82%BC%EC%84%B1%EB%8F%99-359&salesType=one_room%2Ctwo_room%2Cofficetel&tradeType=borrow",
    "난향": "https://www.daangn.com/kr/realty/?areaViewType=p&in=%EB%82%9C%ED%96%A5%EB%8F%99-356&salesType=one_room%2Ctwo_room%2Cofficetel&tradeType=borrow",
    "중앙": "https://www.daangn.com/kr/realty/?areaViewType=p&in=%EC%A4%91%EC%95%99%EB%8F%99-348&salesType=one_room%2Ctwo_room%2Cofficetel&tradeType=borrow",
    "신원":  "https://www.daangn.com/kr/realty/?areaViewType=p&in=%EC%8B%A0%EC%9B%90%EB%8F%99-352&salesType=one_room%2Ctwo_room%2Cofficetel&tradeType=borrow",
    "청림":"https://www.daangn.com/kr/realty/?areaViewType=p&in=%EC%B2%AD%EB%A6%BC%EB%8F%99-342&salesType=one_room%2Ctwo_room%2Cofficetel&tradeType=borrow"
}


# -----------------------------
# 정제 함수
# -----------------------------
def clean_text(text):
    if not isinstance(text, str):
        return text
    text = re.sub(r'[\x00-\x1F\x7F]', '', text)
    text = re.sub(r'[\U00010000-\U0010FFFF]', '', text)
    return text.strip()


# -----------------------------
# 기존 엑셀 불러오기 (수식 무시, 문자열로 읽기)
# -----------------------------
output_file = "daangn_list.xlsx"
if os.path.exists(output_file):
    combined_df = pd.read_excel(output_file, engine="openpyxl", dtype=str)
    print(f"[INFO] 기존 엑셀 불러오기 완료, 현재 행 수: {len(combined_df)}")
else:
    combined_df = pd.DataFrame()
    print(f"[INFO] 기존 엑셀 없음, 신규 생성 예정")

# -----------------------------
# URL 반복 처리
# -----------------------------
for area, url in urls.items():
    print(f"[INFO] {area} 페이지 스크래핑 중...")
    driver.get(url)
    time.sleep(5)

    soup = BeautifulSoup(driver.page_source, "html.parser")
    scripts = soup.find_all("script", type="application/ld+json")
    json_data_list = []

    for script in scripts:
        try:
            data = json.loads(script.string)
            if isinstance(data, dict) and data.get("@type") == "ItemList":
                json_data_list.extend(data.get("itemListElement", []))
        except:
            continue

    print(f"[INFO] {area} 추출 매물 수: {len(json_data_list)}개")

    results = []
    for i, item in enumerate(json_data_list):
        try:
            unit = item.get("item", {})
            description = clean_text(unit.get("description", ""))
            identifier = unit.get("identifier", "")
            images = unit.get("image", [])

            if not identifier:  # identifier 없으면 스킵
                continue

            image_url = ""
            image_count = 0
            if isinstance(images, list) and len(images) > 0:
                image_url = "|".join(images)
                image_count = len(images)

            results.append({
                "area": area,
                "identifier": identifier,
                "description": description,
                "image_count": image_count,
                "image": image_url
            })
        except Exception as e:
            print(f"[WARN] {area} 매물 {i} 추출 오류: {e}")

    # -----------------------------
    # URL별 중복 제거 후 신규 추가
    # -----------------------------
    temp_df = pd.DataFrame(results)
    before_count = len(combined_df)
    combined_df = pd.concat([combined_df, temp_df], ignore_index=True)
    combined_df.drop_duplicates(subset=["identifier"], keep="first", inplace=True)
    added_count = len(combined_df) - before_count
    print(f"[INFO] {area} 신규 추가 매물 수: {added_count}, 누적 행 수: {len(combined_df)}")

# -----------------------------
# 최종 엑셀 저장 (값만 저장, 수식 제거)
# -----------------------------
try:
    combined_df.to_excel(output_file, index=False, engine="openpyxl")
    print(f"[INFO] 최종 엑셀 저장 완료, 총 행 수: {len(combined_df)}")
except Exception as e:
    print(f"[ERROR] 엑셀 저장 실패: {e}")

# -----------------------------
# 종료
# -----------------------------
driver.quit()
print("[INFO] 크롬 드라이버 종료")