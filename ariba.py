import re
import time
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    # browser = playwright.chromium.launch(headless=False,executable_path='C:/Program Files/Google/Chrome/Application/chrome.exe')
    browser = playwright.chromium.launch_persistent_context(headless=False,executable_path="C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe")
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://s1.ariba.com/Sourcing/Main/aw?awh=r&awssk=ODTzCxAIbpbHwlV1&realm=schlumberger")
    # page.goto("https://example.com/")
    # page.get_by_role("heading", name="Example Domain").click()
    # page.get_by_role("link", name="More information...").click()
    # page.get_by_role("link", name="IANA-managed Reserved Domains").click()
    # page.get_by_role("link", name="XN--HLCJ6AYA9ESC7A").click()

    # ---------------------
    time.sleep(15)

    page.fill('#i0116','JQuintero27@slb.com')
    time.sleep(305)

    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
