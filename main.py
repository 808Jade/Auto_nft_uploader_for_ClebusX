import asyncio
import json
import os
import re
import time

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from playwright.async_api import Playwright, async_playwright, Error
import openpyxl

import idpass # for security

def excel_scraping():
    ################################## 각 지점에서 올라온 엑셀파일을 스크래핑 ####################################
    print("== 각 지점에서 올라온 엑셀파일을 스크래핑 ==")

    async def scraping(playwright: Playwright) -> None:
        browser = await playwright.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()
        await page.goto("http://hsmnftupload.almancompany.com/bbs/login.php")
        await page.get_by_placeholder("아이디").click()
        await page.get_by_placeholder("아이디").fill(f"{idpass.admin_id}")
        await page.get_by_placeholder("비밀번호").click()
        await page.get_by_placeholder("비밀번호").fill(f"{idpass.admin_pass}")
        await page.get_by_role("button", name="로그인").click()

        branches = ['용답 ', '수원 ', '인천 ', '성산 ', '대전 ', '원주 ', '수원권선 ', '용인 H']

        for branch in branches:
            await page.get_by_role("link", name=f"{branch}CPO").click()
            first_link = page.locator("li.even a").first
            link_text = await first_link.inner_text()

            if "05" in link_text:
                await first_link.click()

                async with page.expect_download() as download_info:
                    await page.locator('a.view_file_download').click()
                download = await download_info.value
                await download.save_as(f"{branch}.xlsx")
                await page.get_by_role("link", name="그누보드").click()  # 다시 리스트로 돌아감
            else:
                print(f'{branch}지점 업로드 되지 않음.')
                await page.get_by_role("link", name="그누보드").click()  # 다시 리스트로 돌아감

        # ---------------------
        await context.close()
        await browser.close()

    async def main() -> None:
        async with async_playwright() as playwright:
            await scraping(playwright)

    asyncio.run(main())


def excel_to_json():
    ################################## 엑셀을 읽어와서 json파일로 저장 ##########################################
    print("\n== 엑셀을 읽어와서 json파일로 저장 ==")

    branches = ['용답 ', '수원 ', '인천 ', '성산 ', '대전 ', '원주 ', '수원권선 ', '용인 H']

    color_map = {
        "블랙": "Black",
        "그린": "Green",
        "브라운": "Brown",
        "레드": "Red"
        # 필요한 다른 색상 변환 추가
    }

    for branch in branches:
        # 디렉토리 경로 설정
        directory_path = os.path.join(os.getcwd(), branch.strip())

        # 디렉토리 존재 여부 확인 및 생성
        if not os.path.exists(directory_path):
            os.makedirs(directory_path)

        file_path = f'{branch}.xlsx'
        if os.path.exists(file_path):
            df = pd.read_excel(file_path, skiprows=1)
            df['판매일'] = pd.to_datetime(df['판매일'], errors='coerce').dt.strftime('%Y-%m-%d')  # 날짜형식 준수

            # '계약번호'가 null인 행 제거
            df = df.dropna(subset=['계약번호'])

            # 내장색상 열 값 변경 및 없는 색상 출력
            if '내장색상' in df.columns:
                def map_color(color):
                    if pd.isna(color):
                        return "-"
                    elif color in color_map:
                        return color_map[color]
                    else:
                        print(f"    *{branch}파일에서 '{color}' 색상이 color_map에 없습니다.")
                        return color

                df['내장색상'] = df['내장색상'].apply(map_color)

            json_data = df.to_json(orient='records', force_ascii=False)

            json_file_path = os.path.join(directory_path, f'{branch.strip()}.json')
            with open(json_file_path, 'w', encoding='utf-8') as json_file:
                json_file.write(json_data)

            print(f"{branch}파일이 성공적으로 저장되었습니다.")
        else:
            print(f"{branch} PASS")


def make_image():
    ################################## 지점 별 이미지를 생성 ##########################################
    print("\n== 지점 별 이미지를 생성 ==")

    def wrap_text(text, max_length):
        words = text.split()
        lines = []
        current_line = ""

        for word in words:
            if len(current_line) + len(word) + 1 <= max_length:
                if current_line:
                    current_line += " " + word
                else:
                    current_line = word
            else:
                lines.append(current_line)
                current_line = word

        if current_line:
            lines.append(current_line)

        return "\n".join(lines)

    # 글꼴 경로와 크기 설정
    font_path = "MBKCorpoS.otf"  # 글꼴 파일의 경로를 지정
    font_size = 24

    # 계약번호 저장
    nums = {
        '대전': [],
        '성산': [],
        '수원': [],
        '수원권선': [],
        '용답': [],
        '용인 H': [],
        '원주': [],
        '인천': []
    }

    branches = ['용답 ', '수원 ', '인천 ', '성산 ', '대전 ', '원주 ', '수원권선 ', '용인 H']

    for branch in branches:
        # 디렉토리 경로 설정
        directory_path = os.path.join(os.getcwd(), branch.strip())

        # 디렉토리 존재 여부 확인 및 생성
        if not os.path.exists(directory_path):
            os.makedirs(directory_path)

        # 디렉토리 내의 모든 JSON 파일 처리
        for json_filename in os.listdir(directory_path):
            if json_filename.endswith('.json'):
                json_filepath = os.path.join(directory_path, json_filename)

                # JSON 파일 읽어오기
                with open(json_filepath, 'r', encoding='utf-8') as json_file:
                    data = json.load(json_file)

                # 각 항목 처리
                for item in data:
                    # 필요한 데이터 추출 및 문자열로 변환
                    contract_num = item.get('계약번호', '')
                    sale_date = item.get('판매일', '')
                    if isinstance(sale_date, int):  # 타임스탬프일 경우 변환
                        sale_date = pd.to_datetime(sale_date, unit='ms').strftime('%Y-%m-%d')
                    elif isinstance(sale_date, str) and len(sale_date) >= 10:  # 문자열로 날짜 포맷 확인
                        sale_date = pd.to_datetime(sale_date).strftime('%Y-%m-%d')
                    manager = item.get('담당 영업사원', '')
                    year = str(item.get('연식', ''))
                    if isinstance(year, float):
                        year = str(int(year))  # 소숫점 제거
                    else:
                        year = str(year)
                    model = item.get('모델', '')
                    type_of = item.get('차종', '')
                    color_ex = item.get('외부색상', '')
                    color_in = item.get('내장색상', '')
                    warranty_period = item.get('보증기간', '')
                    mileage = str(item.get('판매시 마일리지', '')).replace(" ", "")
                    # 마일리지 형식 변환
                    mileage = re.sub(r'(?<=\d)(?=(\d{3})+(?!\d))', ",", mileage)

                    # 필요한 데이터에 줄바꿈 추가
                    model = wrap_text(model, 20)
                    type_of = wrap_text(type_of, 20)
                    color_ex = wrap_text(color_ex, 20)
                    color_in = wrap_text(color_in, 20)

                    # 이미지 로드
                    img = Image.open('sample.png')

                    # 텍스트 삽입
                    draw = ImageDraw.Draw(img)
                    font = ImageFont.truetype(font_path, font_size)  # 글꼴 크기를 지정

                    positions = [
                        ((52, 1407), contract_num),
                        ((338, 1407), sale_date),
                        ((52, 1506), manager),
                        ((338, 1506), year),
                        ((52, 1604), model),
                        ((338, 1604), type_of),
                        ((52, 1709), color_ex),
                        ((338, 1709), color_in),
                        ((52, 1814), warranty_period),
                        ((338, 1814), mileage)
                    ]

                    # 삽입
                    for position, text in positions:
                        draw.text(position, text, font=font, fill="white")

                    # 이미지 저장
                    nums[branch.strip()].append(contract_num)  # 계약번호 저장

                    output_filename = f"{contract_num}.png"
                    output_filepath = os.path.join(directory_path, output_filename)
                    img.save(output_filepath)
                    print(f"Saved image: {output_filepath}")
    print("Complete.")


################################## ClebusX 에서 발행/보내기 ##########################################
# 특정 지점을 처리할 수 있도록 변수 추가
selected_branch = None


def set_branch(branch_name):
    global selected_branch
    selected_branch = branch_name
    upload_and_send()


def upload_and_send():
    print("\n== ClebusX 에서 발행/보내기 ==")

    data = [
        {'branch': '대전', 'contract_nums': []},
        {'branch': '성산', 'contract_nums': []},
        {'branch': '수원', 'contract_nums': []},
        {'branch': '수원권선', 'contract_nums': []},
        {'branch': '용답', 'contract_nums': []},
        {'branch': '용인 H', 'contract_nums': []},
        {'branch': '원주', 'contract_nums': []},
        {'branch': '인천', 'contract_nums': []},
    ]

    # address.json 파일로부터 계약번호 가져오기
    for num in data:
        json_file_path = os.path.join(num['branch'].strip(), f"{num['branch']}.json")

        if not os.path.exists(json_file_path):
            continue

        with open(json_file_path, 'r', encoding='utf-8') as f:
            d = json.load(f)
            if isinstance(d, list):
                for entry in d:
                    if '계약번호' in entry:
                        num['contract_nums'].append(entry['계약번호'])

    # 선택된 지점만 처리
    branch_data = next((item for item in data if item['branch'] == selected_branch), None)
    if branch_data is None:
        print(f"No data found for branch {selected_branch}")
        return

    async def upload_nft(playwright: Playwright) -> None:
        browser = await playwright.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()
        await page.goto("https://clebusx.com/")
        await page.get_by_role("link", name="MyNFT").click()
        await page.get_by_placeholder("ID").click()
        await page.get_by_placeholder("ID").fill(f"{idpass.hansung_id}")
        await page.get_by_placeholder("ID").press("Tab")
        await page.get_by_placeholder("Password").fill(f"{idpass.hansung_pass}")
        await page.get_by_placeholder("Password").press("Enter")

        if os.path.exists('address.json'):
            with open('address.json', 'r', encoding='utf-8') as f:
                addresses = json.load(f)

        # 발행과 동시에 보내기
        for contract_num in branch_data['contract_nums']:

            # 이미 성공적으로 발행된 경우 건너뛰기
            skip = False
            if isinstance(addresses, list):
                for entry in addresses:
                    if entry['contract_num'] == contract_num and entry.get('Success') == True:
                        print(f"{contract_num} 이미 성공적으로 발행됨. 건너뜀.")
                        skip = True
                        break
            if skip:
                continue

            max_retries = 3
            retries = 0
            success = False

            while retries < max_retries and not success:
                await page.get_by_role("link", name="Sales").click()
                await page.get_by_label("Select the NFT you want to").set_input_files(
                    f"{os.getcwd()}/{branch_data['branch'].strip()}/{contract_num}.png")
                await page.get_by_placeholder("Please enter the NFT name.").click()
                await page.get_by_placeholder("Please enter the NFT name.").fill(
                    f"Mercedes-Benz Certified Warranty Program #{contract_num}")
                await page.get_by_placeholder("Please enter NFT description.").click()
                await page.get_by_placeholder("Please enter NFT description.").fill(
                    "Your vehicle is an official imported vehicle of Mercedes-Benz Korea Co., Ltd., and proves that it is a certified used vehicle that has passed 198 strict quality inspections and inspections of Mercedes-Benz certified used vehicles.\n")
                await page.get_by_role("link", name="Art").click()
                await page.get_by_role("link", name="Automobiles").click()
                await page.get_by_text("Not sold").click()
                await page.get_by_role("button", name="Registration").click()
                time.sleep(2)

                if await page.locator(".market_subject.ell1").count() == 3:
                    print(f"{branch_data['branch']} : {contract_num} 발행되었습니다.")
                    success = True
                else:
                    retries += 1
                    print(
                        f"{branch_data['branch']} : {contract_num} Upload Fail. Retrying... ({retries}/{max_retries})")
                    await asyncio.sleep(2)  # 재시도 전에 약간의 대기 시간을 추가

            if not success:
                print(f"{branch_data['branch']} : {contract_num} 발행을 {max_retries}번 시도했지만 실패하였습니다.")
                continue  # 발행 실패 시 다음 계약번호로 넘어감

            # 보내기
            # address.json 파일로부터 address와 id를 가져옴. contract_num 과 짝을 이룸을 확인함과 동시에 전송
            await page.locator(".chool_a_bt").first.click()

            text = await page.locator(".market_subject.ell1").first.text_content()
            if contract_num in text:
                if os.path.exists('address.json'):
                    with open('address.json', 'r', encoding='utf-8') as f:
                        addresses = json.load(f)

                    if isinstance(addresses, list):
                        for entry in addresses:
                            if entry['contract_num'] == contract_num:
                                address = entry['address']
                                await page.get_by_role("textbox").nth(1).fill(f"{address}")
                                # 재시도 로직 추가
                                max_retries = 3
                                retries = 0
                                success = False

                                while retries < max_retries and not success:
                                    await page.get_by_role("button", name="Send").click()
                                    await asyncio.sleep(1)

                                    if await page.locator(".market_subject.ell1").count() == 2:
                                        success = True
                                        print(f"[{contract_num}] -> {address}")

                                        # 성공적으로 보낸 경우 address.json 파일 업데이트
                                        for entry in addresses:
                                            if entry['contract_num'] == contract_num:
                                                entry['Success'] = True
                                                break
                                        with open('address.json', 'w', encoding='utf-8') as f:
                                            json.dump(addresses, f, ensure_ascii=False, indent=4)

                                    else:
                                        retries += 1
                                        print(f"Send Retrying... ({retries}/{max_retries})")
                                        await asyncio.sleep(2)
                                        try:
                                            await page.locator(".chool_a_bt").first.click()
                                        except Error:
                                            if await page.locator(".market_subject.ell1").count() == 2:
                                                success = True
                                                print(f"[{contract_num}] -> {address}")
                                                for entry in addresses:
                                                    if entry['contract_num'] == contract_num:
                                                        entry['Success'] = True
                                                        break
                                                with open('address.json', 'w', encoding='utf-8') as f:
                                                    json.dump(addresses, f, ensure_ascii=False, indent=4)
                                                break
                                        await page.get_by_role("textbox").nth(1).fill(f"{address}")

                                if not success:
                                    print(f"Failed to send {contract_num} after {max_retries} attempts.")
                                break
                else:
                    print("There is no address.json file.")
                    return

        # ---------------------
        await context.close()
        await browser.close()

    async def main() -> None:
        async with async_playwright() as playwright:
            await upload_nft(playwright)

    asyncio.run(main())


#########################################################################################
from address_saver import save_address_json
import tkinter as tk

root = tk.Tk()
root.title('NFT UPLOADER')
root.geometry("480x640")

button1 = tk.Button(root, text="Excel Scraping", command=excel_scraping)
button2 = tk.Button(root, text="Excel to JSON", command=excel_to_json)
button3 = tk.Button(root, text="Make Image", command=make_image)
button4 = tk.Button(root, text="make 'address.json'", command=save_address_json)

# 버튼 배치
button1.pack(pady=10)
button2.pack(pady=10)
button3.pack(pady=10)
button4.pack(pady=10)

# 지점별 버튼 생성 및 배치
branches = [
    '대전', '성산', '수원', '수원권선', '용답', '용인 H', '원주', '인천'
]

for branch in branches:
    btn = tk.Button(root, text=f"Upload and Send {branch}", command=lambda b=branch: set_branch(b))
    btn.pack(pady=5)

root.mainloop()
#########################################################################################
