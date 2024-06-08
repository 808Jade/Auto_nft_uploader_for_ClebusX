import asyncio

from playwright.async_api import Playwright, async_playwright

import idpass

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

    start_num = 903
    for i in range(100):
        current_num = start_num + i
        await page.get_by_role("link", name="회원추가").click()
        await page.get_by_label("아이디필수").click()
        #################################################################
        await page.get_by_label("아이디필수").fill(f"sg{current_num + 1}a")
        #################################################################
        await page.get_by_label("비밀번호필수").click()
        await page.get_by_label("비밀번호필수").fill(f"{idpass.member_pass}")
        await page.get_by_role("button", name="확인").click()
        await page.get_by_role("link", name="목록").click()

    # ---------------------
    await context.close()
    await browser.close()


async def main() -> None:
    async with async_playwright() as playwright:
        await run(playwright)


asyncio.run(main())
