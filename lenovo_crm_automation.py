from playwright.sync_api import sync_playwright
import time

def run():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(viewport={'width': 1280, 'height': 800})
        page = context.new_page()
        
        print("Открываю SAP CRM...")
        page.goto('https://crm.lenovo.com/sap/crm_logon/default.htm')
        
        print("Жду вашей авторизации (у вас есть 1 минута, чтобы залогиниться)...")
        try:
            # Даем 60 секунд на логин. Как только появится текст роли, скрипт поедет дальше.
            page.wait_for_selector('text="WE - Indirect Sales Rep"', timeout=60000)
            page.locator('text="WE - Indirect Sales Rep"').click()
            print("Роль 'WE - Indirect Sales Rep' выбрана.")
            
            page.wait_for_selector('text="Marketing"', timeout=30000)
            page.locator('text="Marketing"').click()
            print("Вкладка 'Marketing' открыта.")
            
            page.wait_for_selector('text="Trade Promotions"', timeout=10000)
            page.locator('text="Trade Promotions"').click()
            print("Выбраны 'Trade Promotions'.")
            
            # В SAP WebUI кнопка New может быть либо ссылкой, либо спаном
            page.wait_for_selector('text="New"', timeout=30000)
            page.locator('text="New"').first.click()
            print("Кнопка 'New' нажата. Шаги завершены!")
            
        except Exception as e:
            print(f"Что-то пошло не так или истекло время ожидания: {e}")
            
        print("Оставляю браузер открытым на 5 минут, чтобы вы могли продолжить работу...")
        time.sleep(300)
        browser.close()

if __name__ == '__main__':
    run()
