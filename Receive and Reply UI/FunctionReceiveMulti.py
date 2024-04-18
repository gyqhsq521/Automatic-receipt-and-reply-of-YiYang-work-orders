from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from Getinfo import *
import time
import threading
import os

from selenium.webdriver.firefox.options import Options

def Open_Page(driver,username ,password,frequence,district,count,start_time,e3):

    #页面等待加载（显式等待）直到指定元素5出现
    #driver是webdriver对象，10是最长等待时间，0.5是每0.5秒去查询对应的元素。until后面跟的等待具体条件，EC是判断条件，检查元素是否存在于页面的 DOM 上。
    username_input=WebDriverWait(driver,10,0.5).until(EC.presence_of_element_located((By.XPATH,'//*[@id="j_username"]')))

    #输入账号密码
    password_input = driver.find_element(By.XPATH,'//*[@id="j_password"]')
    username_input.clear()
    time.sleep(1)
    username_input.send_keys(username)
    time.sleep(1)
    password_input.send_keys(password)
    time.sleep(1)

    #点击登录按钮
    login_button = driver.find_element(By.XPATH,'/html/body/div/div/div/form/div/div[2]/div/ul/li[3]/input[1]')
    login_button.click()

    #页面等待加载（显示等待（有条件等待））直到新页面的iframe出现,并将iframe作为可操作对象
    iframe=WebDriverWait(driver,10,0.5).until(EC.presence_of_element_located((By.XPATH,'//*[@id="mainFrame"]')))
    driver.switch_to.frame(iframe)

    #等待“陕西联通集中故障处理流程”元素出现，并双击
    directory_sxlt=WebDriverWait(driver,10,0.5).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[2]/div[2]/div/div/ul/li/ul/li[1]/div/a/span')))
    ActionChains(driver).double_click(directory_sxlt).perform()

    #等待“待处理”出现，并单击
    directory_dcl=WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[2]/div[2]/div/div/ul/li/ul/li[1]/ul/li[2]/div/a/span')))
    directory_dcl.click()


    #将iframe1作为可操作对象
    driver.switch_to.frame(driver.find_element(By.XPATH,'//*[@id="pages"]'))

    


    #等待表格元素出现
    try:
        WebDriverWait(driver,2,0.5).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[4]/div/div/table/tbody/tr[1]/td[3]')))
    except Exception:
        count += 1    
        print(f"{datetime.now().replace(microsecond=0)}：{district}：第{count}次接单，用时{(datetime.now() - start_time).seconds}秒，该账号暂时没有工单")              
        driver.quit()
        e3 += 1
        time.sleep(int(frequence)*60)

    return e3


def Recieve_gongdan(username ,password,frequence,district,mode_type):
    count = 0
    #设置无浏览器窗口运行
    f_options = Options()
    f_options.add_argument('--headless')
    
    while True:

        #用来确认当前账号是否有工单的信号值
        e3 = 0

        try:
            start_time = datetime.now()

            if mode_type == "是":
                #参数填为无浏览器窗口运行
                driver = webdriver.Firefox(options = f_options)
            else:
                driver = webdriver.Firefox()
                
            url="http://10.93.19.175:8091/wyeoms/"

            #打开指定网址的网页
            driver.get(url)

            try:
                #Open_Page函数执行打开工单页面操作，并将函数返回结果赋值给e3
                e3 = Open_Page(driver,username ,password,frequence,district,count,start_time,e3)
                if e3 == 1:
                    continue
            except Exception as e4:
                print(f"{datetime.now().replace(microsecond=0)}：{district}：错误4",e4)
                try:
                    driver.quit()
                    continue
                except Exception as e5:
                    print(f"{datetime.now().replace(microsecond=0)}：{district}：窗口已关闭：错误5：",e5)
                    continue

            data_list_name = []
            data_list_url = []

            try:
                #点击“最后一页”
                page_turn0 = driver.find_element(By.LINK_TEXT,'最后一页')
                page_turn0.click()

                #获取当前账号工单总页数
                page_number = get_pagenum(get_html_str(driver.page_source))
                
                #回到“第一页”
                page_turn1 = driver.find_element(By.LINK_TEXT,'第一页')
                page_turn1.click()
            except Exception as eB:
                page_number = 1
            
            #爬取网页表格数据，page_number为几，就爬到第几页
            for k in range(0,int(page_number)):

                #爬取第十一列（任务所有者）page_source为当前页面源码
                data_list_name = data_list_name + get_data(get_html_str(driver.page_source))

                #爬取表格第三列（工单流水号）链接
                data_list_url = data_list_url + get_url(get_html_str(driver.page_source))

                if int(page_number)-1 > k >= 0:
                    #点击"下一页"
                    page_turn = driver.find_element(By.LINK_TEXT,'下一页')
                    page_turn.click()
                    
                    #等待表格元素出现
                    WebDriverWait(driver,10,0.5).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[4]/div/div/table/tbody/tr[1]/td[3]')))

       
            #获取爬取列表的元素数量
            current_length = len(data_list_url)
            
           
            #接工单
            j = 0
            for i in range(1,current_length+1):
                if data_list_name[i-1] == '':

                    j = j+1
                    
                    #在新标签页打开未接单的工单
                    driver.switch_to.new_window('tab')
                    driver.get("http://10.93.19.175:8091/wyeoms/sheet/centralfaultprocess/{}".format(data_list_url[i-1]))
    ##                driver.execute_script("window.open('{}','_blank');".format(data_list_url[i-1])

                    try:
                        #等待“原因初判”下拉列表元素出现
                        selectTag = WebDriverWait(driver,10,0.5).until(EC.presence_of_element_located((By.XPATH,'//*[@id="remark"]')))

                        #目前西乡接单、回单需要填操作联系人方式
                        if district == "西乡":
                            b = driver.find_element(By.XPATH,'//*[@id="operaterContact"]')
                            b.send_keys("15619161056")
                            time.sleep(0.2)
                        
                        # 实例化一个select对象：传入Select标签元素的Element对象
                        select = Select(selectTag)

                        # 选择“动力环境故障”       
                        select.select_by_visible_text('动力环境故障')
                        time.sleep(1)
                    except Exception:
                        continue

                    #点击“确认受理”按钮
                    confirm_button = driver.find_element(By.XPATH,'/html/body/div[4]/div/div/div[2]/div[2]/div/form/div[3]/input')
                    confirm_button.click()

                    #等待受理完成
                    try:
                        complete = WebDriverWait(driver,10,0.5).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[4]/div/div/div[2]/div[2]/div/form/div[2]/fieldset/div/div[1]/div/div/div')))
                    except Exception:
                        driver.close()
                        driver.switch_to.window(driver.window_handles[-1])  
                        continue
                    
                    #关闭标签页
                    driver.close()

                    #切回当前标签页，使当前标签页作为可操作对象
                    driver.switch_to.window(driver.window_handles[-1])                    

            #关闭浏览器窗口
            driver.quit()

            end_time = datetime.now()
            count += 1
            print(f"{datetime.now().replace(microsecond=0)}：{district}：第{count}次接单，用时{(end_time - start_time).seconds}秒,本次共有{j}个待接工单")
            time.sleep(int(frequence)*60)
        except Exception as e1:
            print(f"{datetime.now().replace(microsecond=0)}：{district} 错误1：",e1)
            try:
                driver.quit()
            except Exception as e2:
                print(f"{datetime.now().replace(microsecond=0)}：{district} 关闭窗口错误：错误2：",e2)


#清除后台卡死进程
def clear(f):
    while True:
        os.system('taskkill /F /IM firefox.exe')
        os.system('taskkill /F /IM geckodriver.exe')
        time.sleep(int(f)*60)

#创建线程实例，方便在主函数中多线程创建
def t(u,p,f,d,m):
    a = threading.Thread(target=Recieve_gongdan,args=(u,p,f,d,m))
    a.daemon = True
    a.start()
    
#创建清除后台的线程
def c(f):
    a = threading.Thread(target=clear,args=(f,))
    a.daemon = True
    a.start()



    



         

