import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import json
import re

# -----------------------------
# 1. 엑셀 읽기
# -----------------------------
input_file = "daangn_list.xlsx"
df = pd.read_excel(input_file, engine="openpyxl", dtype=str)

# -----------------------------
# 2. 크롬 드라이버 설정
# -----------------------------
driver_path = "C:/Users/USER/Downloads/chromedriver-win64/chromedriver.exe"
options = Options()
options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")
options.add_argument("--log-level=3")
service = Service(driver_path)
driver = webdriver.Chrome(service=service, options=options)


# -----------------------------
# 3. 상세 스크래핑 함수(WebDriverWait 적용)
# -----------------------------
def scrape_detail(url):
    driver.get(url)

    try:
        # dt 태그가 나타날 때까지 최대 10초 대기
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "dt"))
        )
    except:
        print(f"[WARN] {url} 로딩 실패")
        return {}

    soup = BeautifulSoup(driver.page_source, "html.parser")

    # dt/dd 태그 매칭
    dt_tags = [dt.get_text(strip=True) for dt in soup.find_all("dt")]
    dd_tags = [dd.get_text(strip=True) for dd in soup.find_all("dd")]
    info_dict = {dt: dd for dt, dd in zip(dt_tags, dd_tags)}

    # 필요한 컬럼만 필터링
    keep_keys = ["아파트명", "건축물 용도", "전용면적", "층", "방향", "관리비", "사용승인일 (연식)"]
    filtered_info = {k: info_dict.get(k, None) for k in keep_keys}

    # -----------------------------
    # 정제
    # -----------------------------
    floor_info = filtered_info.get("층", "")
    if isinstance(floor_info, str):
        match = re.search(r"(\d+)층\s*/\s*(\d+)층", floor_info)
        if match:
            filtered_info["층"] = int(match.group(1))
            filtered_info["총층"] = int(match.group(2))
        else:
            single_match = re.search(r"(\d+)층", floor_info)
            filtered_info["층"] = int(single_match.group(1)) if single_match else None
            filtered_info["총층"] = None
    else:
        filtered_info["층"] = None
        filtered_info["총층"] = None

    direction = filtered_info.get("방향", "")
    filtered_info["방향"] = re.sub(r"\s*\(.*?\)", "", direction).strip() if direction else None

    fee = filtered_info.get("관리비", "")
    if fee:
        fee_numbers = re.findall(r"(\d+)만원", fee)
        fee_numbers = list(set(fee_numbers))
        filtered_info["관리비"] = sum(int(n) for n in fee_numbers) * 10000 if fee_numbers else None
    else:
        filtered_info["관리비"] = None

    area_raw = filtered_info.get("전용면적", "")
    if area_raw:
        area_match = re.search(r"([\d.]+)㎡", area_raw)
        filtered_info["전용면적"] = float(area_match.group(1)) if area_match else None
    else:
        filtered_info["전용면적"] = None

    date_raw = filtered_info.get("사용승인일 (연식)", "")
    if date_raw:
        date_match = re.search(r"(\d{4})년\s*(\d{2})월\s*(\d{2})일", date_raw)
        filtered_info[
            "사용승인일 (연식)"] = f"{date_match.group(1)}-{date_match.group(2)}-{date_match.group(3)}" if date_match else None
    else:
        filtered_info["사용승인일 (연식)"] = None

    scripts = soup.find_all("script", {"type": "application/ld+json"})
    price, street_address = None, None
    for script in scripts:
        try:
            data = json.loads(script.string.strip())
        except Exception:
            continue

        if data.get("@type") == "SingleFamilyResidence":
            addr = data.get("address", {})
            street_address = addr.get("streetAddress", street_address)
        elif data.get("@type") == "AggregateOffer":
            offers = data.get("offers", [])
            if isinstance(offers, list) and len(offers) > 0:
                price = offers[0].get("price")

    time_tag = soup.find("time")
    date_value = time_tag["datetime"].split("T")[0] if time_tag and time_tag.has_attr("datetime") else None

    filtered_info.update({
        "가격": price,
        "주소": street_address,
        "등록일": date_value
    })

    return filtered_info


# -----------------------------
# 4. 각 행 반복하여 상세 정보 추가
# -----------------------------
for idx, row in df.iterrows():
    url = row.get("identifier")
    if pd.isna(url):
        continue
    print(f"[INFO] {idx + 1}/{len(df)} 상세 페이지 스크래핑: {url}")
    detail_data = scrape_detail(url)
    for k, v in detail_data.items():
        df.at[idx, k] = v

# -----------------------------
# 5. 엑셀 저장
# -----------------------------
output_file = "daangn_list_detail.xlsx"
df.to_excel(output_file, index=False, engine="openpyxl")
print(f"[INFO] 최종 엑셀 저장 완료: {output_file}")

driver.quit()
print("[INFO] 크롬 드라이버 종료")
