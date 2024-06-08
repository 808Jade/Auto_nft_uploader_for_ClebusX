import asyncio
import json
import os

from playwright.async_api import Playwright, async_playwright

import idpass

# address_saver.py
# 지점별로 dir 속에 있는 json파일로부터 필요한 만큼의 'id & address & contract_num' 을 'address.json' 으로 저장함.
#   * 각 지점 별 start_id는 사용자가 임의로 작성해야 함.

def save_address_json():
    async def run(playwright: Playwright) -> None:
        browser = await playwright.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()
        await page.goto("https://clebusx.com")

        await page.get_by_role("link", name="LOGIN").click()
        await page.get_by_placeholder("ID").click()
        await page.get_by_placeholder("ID").fill(f"{idpass.admin_id}")
        await page.get_by_placeholder("ID").press("Tab")
        await page.get_by_placeholder("Password").click()
        await page.get_by_placeholder("Password").fill(f"{idpass.admin_pass}")
        await page.get_by_role("button", name="LOGIN").click()

        await page.goto("https://clebusx.com/adm")

        await page.get_by_role("button", name="회원관리").click()
        await page.get_by_role("link", name="회원관리", exact=True).click()

        # 대전(djn)    620~
        # 성산(ssa)    621~
        # 수원(swn)    1051~
        # 수원권선(sg)  903~
        # 용답(ydp)    1233~
        # 용인 H(yinh) 288~
        # 원주(wju)    441~
        # 인천(icn)    502~ (502 까지 존재)

        branches = [
            {'name_for_dir': '대전 ', 'name': 'djn', 'start_num': 620, 'repeat': 0, 'contract_nums': []},
            {'name_for_dir': '성산 ', 'name': 'ssa', 'start_num': 621, 'repeat': 0, 'contract_nums': []},
            {'name_for_dir': '수원 ', 'name': 'swn', 'start_num': 1051, 'repeat': 0, 'contract_nums': []},
            {'name_for_dir': '수원권선 ', 'name': 'sg', 'start_num': 903, 'repeat': 0, 'contract_nums': []},
            {'name_for_dir': '용답 ', 'name': 'ydp', 'start_num': 1233, 'repeat': 0, 'contract_nums': []},
            {'name_for_dir': '용인 H', 'name': 'yinh', 'start_num': 288, 'repeat': 0, 'contract_nums': []},
            {'name_for_dir': '원주 ', 'name': 'wju', 'start_num': 441, 'repeat': 0, 'contract_nums': []},
            {'name_for_dir': '인천 ', 'name': 'icn', 'start_num': 502, 'repeat': 0, 'contract_nums': []}
        ]

        extracted_data = []

        for branch in branches:
            json_file_path = os.path.join(branch['name_for_dir'].strip(), f"{branch['name_for_dir'].strip()}.json")

            if not os.path.exists(json_file_path):
                continue

            with open(json_file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, list):
                    branch['repeat'] = len(data)
                    for entry in data:
                        if '계약번호' in entry:
                            branch['contract_nums'].append(entry['계약번호'])

        for branch in branches:
            for i in range(branch['repeat']):
                current_num = branch['start_num'] + i

                await page.get_by_label("검색어 필수").click()
                await page.get_by_label("검색어 필수").fill(f"{branch['name']}{current_num + 1}a")
                await page.get_by_label("검색어 필수").press("Enter")

                text = await page.locator('td.td_date a', has_text="0x").text_content()
                if text:  # text가 None이 아닌 경우만 처리
                    extracted_data.append({
                        'id': f"{branch['name']}{current_num + 1}a",
                        'address': text.strip(),
                        'contract_num': branch['contract_nums'][i] if i < len(branch['contract_nums']) else None
                    })

                    print(f"{branch['name']}{current_num + 1}a [{branch['contract_nums'][i]}] = {text}")
            print('\n')

        # JSON 파일로 저장
        output_file_path = os.path.join(os.getcwd(), 'address.json')
        with open(output_file_path, 'w', encoding='utf-8') as json_file:
            json.dump(extracted_data, json_file, ensure_ascii=False, indent=4)

        # ---------------------
        await context.close()
        await browser.close()

    async def main() -> None:
        async with async_playwright() as playwright:
            await run(playwright)

    asyncio.run(main())
