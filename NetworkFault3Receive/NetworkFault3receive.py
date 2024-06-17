from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import datetime
import time
from getinfo import *
import os

import threading
from pynput import mouse,keyboard

def receive_order(driver,j):
   
    #获取工单列表第8列的数据
    data_list = []
    data_list +=  get_processing_stage(get_html_str(driver.page_source))           

    #待办工单数量
    c_len = len(data_list)
    
    table = driver.find_element(By.XPATH,'//*[@id="app"]/section/section/main/div[1]/div[2]/div/div/div[2]/div/div[3]/table/tbody')
    table_rows = table.find_elements(By.TAG_NAME,'tr')
    #点击待受理工单
    for i in range(0,c_len):
        if data_list[i] == "受理":  
            #点击进入待受理工单
            order = table_rows[i].find_element(By.XPATH,'./td[3]/div/button')
            order.click()

            #点击“受理”
            try:
                rec = WebDriverWait(driver,10,0.5).until(EC.presence_of_element_located((By.XPATH,'//*[@id="app"]/section/section/main/div[1]/div[1]/div[1]/div[2]/button')))
                rec.click()
                
            except Exception as e5:
                print(f"接单失败:\n{e5}")

   
            #关闭工单详情页(最后一个详情页）
            time.sleep(3)
            close = WebDriverWait(driver,10,0.5).until(EC.presence_of_element_located((By.XPATH,'//*[@id="app"]/section/section/div/div/div/div[1]/div/span[last()]/span')))
            close.click()
##            driver.execute_script("arguments[0].click;",close)

            j += 1
            
            try:
                #等待待处理工单列表加载完成
                WebDriverWait(driver,10,0.5).until(EC.presence_of_element_located((By.XPATH,'//*[@id="app"]/section/section/main/div[1]/div[2]/div/div/div[2]/div/div[3]/table/tbody/tr[1]/td[3]')))
            except Exception as e4:
                print("接单已完成，未发现待处理工单")
                
    return j

def MouseMov():
    m = mouse.Controller()
    k = keyboard.Controller()
    while True:
        m.position = (200,200)
        m.click(mouse.Button.right,1)
        time.sleep(0.1)
        m.position = (250,200)
        m.click(mouse.Button.right,1)
        time.sleep(0.1)
        k.press(keyboard.Key.esc)
        time.sleep(420)

def Receive(frequence,pid,mode_type):
##if __name__ == '__main__':
##    os.popen(f'start chrome --remote-debugging-port=9222 --user-data-\dir="D:\selenium" http://sso.portal.unicom.local/eip_sso/aiportalLogin.html?appid=na186&oawx_t=A0002')
##    a = input("输入回车后开始脚本...")
    
##    c = threading.Thread(target=MouseMov)
##    c.daemon = True
##    c.start()

    f_options = webdriver.ChromeOptions()   
    #连接先前新创建的远程调试模式的fiefox
    f_options.debugger_address = f"127.0.0.1:{pid}"
##    f_options.debugger_address = f"127.0.0.1:9222"

    #无浏览器窗口运行
##    if mode_type == '是':
##        f_options.add_argument('headless')

    driver = webdriver.Chrome(options = f_options) 

    count = 0
    
    #切至故障工单标签页，使当前标签页作为可操作对象
    driver.switch_to.window(driver.window_handles[-1])

    #获取当前标签页的句柄
    c_w = driver.current_window_handle
    while True:
        try:         
            start_time = datetime.datetime.now()
            j = 0
            try:

                #切至当前标签页，使当前标签页作为可操作对象
                driver.switch_to.window(c_w)
                print("当前页面:",driver.title)

                #等待“待办工单”加载完成
                wo = WebDriverWait(driver,10,0.3).until(EC.presence_of_element_located((By.XPATH,'//*[@id="app"]/section/aside/div[2]/div[1]/div/ul/div[3]/a')))
                wo.click()
       
                #等待“刷新”按钮加载完成
                refresh = WebDriverWait(driver,10,0.3).until(EC.presence_of_element_located((By.XPATH,'//*[@id="app"]/section/section/main/div[1]/div[2]/div/div/div[1]/div/div[2]/div/button[last()]')))
                refresh.click()
                                                                  
                try:
                    #等待/确定工单列表加载完成                                                           
                    WebDriverWait(driver,10,0.3).until(EC.presence_of_element_located((By.XPATH,'//*[@id="app"]/section/section/main/div[1]/div[2]/div/div/div[2]/div/div[3]/table/tbody/tr[1]')))

                except Exception as e6:
                    end_time = datetime.datetime.now()
                    count += 1
                    print(f"{datetime.datetime.now().replace(microsecond=0)}：第{count}次接单，用时{(end_time - start_time).seconds}秒,本次未发现待办工单")
                    time.sleep(float(frequence)*60)
                    continue
                
            except Exception as e3:
                print(f"{datetime.datetime.now().replace(microsecond=0)}：错误3：",e3)
                count += 1
                continue

            try:  
                #开始受理流程,函数执行完成后返回待受理工单数量
                k = receive_order(driver,j)

            except Exception as e2:
                print(f"{datetime.datetime.now().replace(microsecond=0)}：错误2：",e2)
                count += 1
                continue
            
            count += 1
            end_time = datetime.datetime.now()
            print(f"{datetime.datetime.now().replace(microsecond=0)}：第{count}次接单，用时{(end_time - start_time).seconds}秒,本次共有{k}个待接工单")
            time.sleep(float(frequence)*60)
            
        except Exception as e1:
            count += 1
            print(f"{datetime.datetime.now().replace(microsecond=0)}：错误1：",e1)

        
