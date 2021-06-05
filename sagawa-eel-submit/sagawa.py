from selenium import webdriver
from selenium.webdriver import Chrome, ChromeOptions
from time import sleep
import pandas as pd
from datetime import  datetime
from webdriver_manager.chrome import ChromeDriverManager
import eel
import openpyxl

def set_driver(driver_path,headless_flg):
    options = ChromeOptions()

    if headless_flg == True:
        options.add_argument('--headless')
    
    options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.116 Safari/537.36')
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--ignore-ssl-errors')
    options.add_argument('--incognito')  # シークレットモードの設定を付与

    return Chrome(ChromeDriverManager().install(),options=options)

time=datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
log_file_name = 'log/log_{}.log'.format(time)

def log(txt):
    now = datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
    logStr = '[log: {}] {}'

    with open(log_file_name,'a',encoding='utf-8_sig') as f:
        f.write(logStr.format(now,txt)+'\n')
    print(logStr.format(now,txt))
    eel.view_log(logStr.format(now,txt))

# url = 'https://k2k.sagawa-exp.co.jp/p/sagawa/web/okurijoinput.jsp'

def main(xlname):
    url = 'https://k2k.sagawa-exp.co.jp/p/sagawa/web/okurijoinput.jsp'
    if xlname == '':
        print('ファイル名を入力してください。')
        eel.view_log('ファイル名を入力してください。')
    else:
        log('処理開始')
        log('ファイル名：{}、URL：{}'.format(xlname,url))
        df = pd.read_excel(xlname)
        numbers = df['問い合わせ番号']
        driver = set_driver('chromedriver.exe',False)
        driver.get(url)
        sleep(2)
    
    results = []
    count = 0
    success = 0
    fail = 0
    for number in numbers:
        try:
            driver.find_element_by_id('main:no1').send_keys(str(number))
            driver.find_element_by_id('main:toiStart').click()
            result = driver.find_element_by_css_selector('#list1 > div > table > tbody > tr:nth-child(2) > td')
            results.append(result.text)
            log('{}件目確認成功：{}'.format(count,number))
            success += 1
            driver.find_element_by_id('_id24').click()
            sleep(1)
        except Exception as e:
            log(f'{count}件目確認失敗：{number}')
            log(e)
            fail += 1
        finally:
            # finallyは成功の時もエラーの時も関係なく必ず実行する処理。ここでは必ずcountに１をプラスして総合回数を記録する処理。        
            count += 1
    log('結果確認完了しました。確認総数：{}件、確認成功：{}件、確認失敗{}件'.format(count,success,fail))

    driver.quit()

    df['結果'] = results
    df.to_excel(xlname,index=False)

    
    
    
