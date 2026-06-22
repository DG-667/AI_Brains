import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

def run():
    print("Настраиваю WebDriver и запускаю Chrome...")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)
    wait = WebDriverWait(driver, 60)

    try:
        print("Открываю SAP CRM...")
        driver.get('https://crm.lenovo.com/sap/crm_logon/default.htm')
        
        print("Ожидаю ссылку 'WE - Indirect Sales Rep'. Войдите в систему...")
        role_link = wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'WE - Indirect Sales Rep')]")))
        driver.execute_script("arguments[0].click();", role_link)
        print("Нажали на роль. Ждем загрузки тяжелого интерфейса SAP (10 сек)...")
        
        time.sleep(10) # Жесткая пауза, так как SAP часто перерисовывает интерфейс
        
        print("Ищу вкладку Marketing...")
        marketing_link = wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Marketing')]")))
        driver.execute_script("arguments[0].click();", marketing_link)
        print("Кликнули Marketing. Ждем раскрытия меню (3 сек)...")
        
        time.sleep(3)
        
        print("Ищу Trade Promotions...")
        trade_promo_link = wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Trade Promotions')]")))
        driver.execute_script("arguments[0].click();", trade_promo_link)
        print("Кликнули Trade Promotions. Ждем загрузки таблицы (5 сек)...")
        
        time.sleep(5)
        
        print("Ищу кнопку New...")
        new_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'New')]")))
        driver.execute_script("arguments[0].click();", new_btn)
        print("Успех! Кнопка New нажата.")
        
    except Exception as e:
        print(f"\n[ОШИБКА] Скрипт споткнулся. Детали ошибки:\n{e}\n")
    
    print("Оставляю браузер открытым на 5 минут...")
    time.sleep(300)
    driver.quit()

if __name__ == '__main__':
    run()
