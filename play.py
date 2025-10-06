from playwright.sync_api import sync_playwright

def run():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, devtools=True, slow_mo=250)
        page = browser.new_page()
        page.goto("https://s1.ariba.com/Sourcing/Main/aw?awh=r&awssk=ODTzCxAIbpbHwlV1&realm=schlumberger")
        # Aquí podrás usar el Inspector para avanzar paso a paso
        # page.click("css=button")
        browser.close()

if __name__ == "__main__":
    run()
