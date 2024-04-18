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

from FunctionReceiveMulti import Open_Page

5#下拉菜单选择
def reply_select(driver,loc,option):
    select1 = driver.find_element(By.XPATH,loc)
    select = Select(select1)
    select.select_by_visible_text(option)
    time.sleep(0.2)

#回复工单页面各项选项选择
def reply_order(driver,data_list_url,i,Type,district):

    #在新标签页打开要回复的工单
    driver.switch_to.new_window('tab')
    url = "http://10.93.19.175:8091/wyeoms/sheet/centralfaultprocess/{}".format(data_list_url[i-1])
    driver.get(url)

    #若有阶段性回复，则进行阶段性回复
    try:
        jz = WebDriverWait(driver,3,0.3).until(EC.presence_of_element_located((By.XPATH,'//*[@id="progressDesc"]')))

        reply_select(driver,'//*[@id="remark"]','动力环境故障')        
        
        #进展描述填“可回复”        
        jz.send_keys("可回复")
        
        #点击提交
        tj = driver.find_element(By.XPATH,'//*[@id="method.save"]')
        tj.click()

        #等待受理完成
        WebDriverWait(driver,10,0.5).until(EC.presence_of_element_located((By.XPATH,'/html/body/div/div/div/div/h1')))
                   
        #再次打开工单页面    
        driver.get(url)            

    except Exception:
        pass
    
    #等待“故障级别”下拉列表元素出现   
    #实例化一个select对象：传入Select标签元素的Element对象
    #选择“一般”
    select0 = WebDriverWait(driver,3,0.5).until(EC.presence_of_element_located((By.XPATH,'//*[@id="fault_level"]'))) 
    select = Select(select0)                      
    select.select_by_visible_text('一般')
    time.sleep(0.2)

    #目前西乡接单、回单需要填操作联系人方式
    if district == "西乡":
        b = driver.find_element(By.XPATH,'//*[@id="operaterContact"]')
        b.send_keys("15619161056")
        time.sleep(0.2)

    try:
        reply_select(driver,'//*[@id="fault_deal_result"]','已解决')
    except Exception:
        reply_select(driver,'//*[@id="fault_deal_result"]','无需解决')
        
    reply_select(driver,'//*[@id="fault_reason_one"]',Type)
    reply_select(driver,'//*[@id="fault_reason_two"]','供电故障停电')
    reply_select(driver,'//*[@id="deal_measure"]','自动恢复')
    reply_select(driver,'//*[@id="is_final_way"]','是')

    #若需要选择故障消除时间与业务恢复时间进行的以下操作
    try:
        #点击故障消除时间文本框、并点击弹出菜单中的确定
        fc = driver.find_element(By.XPATH,'//*[@id="fault_clean_time"]')
        fc.click()
        fc1 = WebDriverWait(driver,10,0.3).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[1]/table/tbody/tr[4]/td/div/input[2]')))
        fc1.click()

        #点击业务恢复时间文本框、并点击弹出菜单中的确定
        br = driver.find_element(By.XPATH,'//*[@id="business_recover_time"]')
        br.click()
        br1 = WebDriverWait(driver,10,0.3).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[1]/table/tbody/tr[4]/td/div/input[2]')))
        br1.click()
    except Exception:
        pass
        
#选择子告警手动清除时间为当前时间
def reply_subalert(driver):
    #视图上移
    up = driver.find_element(By.XPATH,'//*[@id="sheet"]/tbody/tr[14]/td[2]/textarea')
    up.click()

    #定位到ifame1下面的元素，将操作视图滚动到iframe1的位置（否则无法选择时间）
    down = driver.find_element(By.XPATH,'//*[@id="dealSelector"]')
    down.click()
    
    #在工单回复页面将driver切换到iframe1，以操作iframe1里的元素
    iframe1 = driver.find_element(By.XPATH,'//*[@id="iframe1"]')
    driver.switch_to.frame(iframe1)

    #获取子告警总页数
    an = get_pages(driver)

    for l in range(0,int(an)):
        #定位子告警表格元素
        table = driver.find_element(By.XPATH,'/html/body/div[4]/div/div/table/tbody')

        #tbody包含行数的集合（注意这里的find_element用复数形式)
        table_rows = table.find_elements(By.TAG_NAME,'tr')

        #获取总行数
        vrows = len(table_rows)  

        #将所有子告警手工清除时间设置为当前时间
        for i in range(0,vrows):
            #点击文本框
            sa1 = table_rows[i].find_element(By.XPATH,'./td[11]/input')
            sa1.click()
                
            #点击下拉菜单中的“确定”
            sa2 = WebDriverWait(driver,10,0.3).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[1]/table/tbody/tr[4]/td/div/input[2]'))) 
            sa2.click()

        if int(an)-1 > l >= 0:

            #点击"下一页"
            page_turn = driver.find_element(By.LINK_TEXT,'下一页')
            page_turn.click()

            #等待表格元素出现
            WebDriverWait(driver,10,0.5).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[4]/div/div/table/tbody')))

            #切回主iframe，将将操作视图滚动到iframe1全部显示的位置
            driver.switch_to.default_content()
            down.click()
            driver.switch_to.frame(iframe1)

    #点击保存
    time.sleep(2)
    save = driver.find_element(By.XPATH,'/html/body/div[4]/div/div/div[1]/input')
    save.click()
  
    #将driver切回主页面，下一步操作主页面上的元素
    driver.switch_to.default_content()

#提交回复工单
def reply_submit(driver):

    #点击“提交”按钮
    submit_button = driver.find_element(By.XPATH,'//*[@id="method.save"]')
    submit_button.click()
    time.sleep(1.5)
    try:
        #切换至弹窗
        al = driver.switch_to.alert

        #点击弹窗的确定
        al.accept()
        
    except Exception:
        pass

    #等待受理完成
    WebDriverWait(driver,10,0.5).until(EC.presence_of_element_located((By.XPATH,'/html/body/div/div/div/div/h1')))

    #关闭标签页
    driver.close()
    
    #切回当前标签页，使当前标签页作为可操作对象
    driver.switch_to.window(driver.window_handles[-1])

#获取总页数
def get_pages(driver):   
    try:
        #点击“最后一页”
        page_turn0 = driver.find_element(By.LINK_TEXT,'最后一页')
        page_turn0.click()

        #获取当前账号工单总页数
        page_number = get_pagenum(get_html_str(driver.page_source))
        
        #回到“第一页”
        page_turn1 = driver.find_element(By.LINK_TEXT,'第一页')
        page_turn1.click()
    except Exception:
        page_number = 1

    return page_number


def reply_gongdan(username,password,frequence,district,mode_type):
    count = 0

    #设置无浏览器窗口运行
    f_options = Options()
    f_options.add_argument('--headless')

    while True:

        #用来确认当前账号是否有工单的信号值
        e3 = 0

        #用来确认有几个工单未成功回复的信号值
        ur = 0
        
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
                print(f"{datetime.now().replace(microsecond=0)}：{district}：错误4：",e4)
                try:
                    driver.quit()
                    continue
                except Exception as e5:
                    print(f"{datetime.now().replace(microsecond=0)}：{district} 关闭窗口错误：",e5)
                    continue
                
            data_list_name = []
            data_list_alarm = []
            data_list_type = []
            data_list_time = []
            data_list_url = []

            #获取工单总页数
            pn = get_pages(driver)

            #爬取网页表格数据，page_number为几，就爬到第几页
            for k in range(0,int(pn)):
                
                #爬取第十一列（任务所有者）page_source为当前页面源码
                data_list_name = data_list_name + get_data(get_html_str(driver.page_source))

                #获取第14列(告警清除时间)的值
                data_list_alarm += get_alarm(get_html_str(driver.page_source))

                #获取第8列（专业）的值
                data_list_type += get_type(get_html_str(driver.page_source))

                #获取第6列（派单时间）的值
                data_list_time += get_time(get_html_str(driver.page_source))
                
                #获取第3列（工单流水号）链接
                data_list_url += get_url(get_html_str(driver.page_source))

                if int(pn)-1 > k >= 0:
                    #点击"下一页"
                    page_turn = driver.find_element(By.LINK_TEXT,'下一页')
                    page_turn.click()
                    
                    #等待表格元素出现
                    WebDriverWait(driver,10,0.5).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[4]/div/div/table/tbody/tr[1]/td[3]')))             

            #获取爬取列表的元素数量
            current_length = len(data_list_url)
                  
            #接工单
            j = 0

            #无线网工单的故障原因一级分类回复选项
            typeA = '配套'

            #传送网工单的故障原因一级分类回复选项
            typeB = '线路故障'

            #接入网工单的故障原因一级分类回复选项
            typeC = '动力配套故障'

            #互联网工单的故障原因一级分类回复选项
            typeD = '电源及配套设施原因'

            #动环网工单的故障原因一级分类回复选项
            typeE = '开关电源'
            
            for i in range(1,current_length+1):
                #确认回单前当前时间
                c_time = datetime.now().replace(microsecond=0)

                if data_list_name[i-1] != '':
                
                    #回复超过派单时间2个小时的工单
                    if (c_time - data_list_time[i-1]).total_seconds() >= 7200.0:        
                        try:
                            #回复无线网工单
                            if data_list_type[i-1] == '无线网':
                                j += 1

                                #开始回复工单
                                reply_order(driver,data_list_url,i,typeA,district)

                                reply_select(driver,'//*[@id="maintenanceSubject"]','联通')
                                reply_select(driver,'//*[@id="mobileFaultDutyId"]','联通')

                                try:
                                    #选择子告警手动清除时间为当前时间
                                    reply_subalert(driver)
                                except Exception:
                                    pass   

                                #将driver切回主页面，下一步操作主页面上的元素
                                driver.switch_to.default_content()
                                
                                #提交并等待完成
                                reply_submit(driver)

                            #回复传送网工单
                            if data_list_type[i-1] == '传输':
                                j += 1

                                #开始回复工单
                                reply_order(driver,data_list_url,i,typeB,district)

                                reply_select(driver,'//*[@id="maintenanceSubject"]','联通')                    
                                reply_select(driver,'//*[@id="transLineLevel"]','本地')
                                reply_select(driver,'//*[@id="faultStatus"]','处理结束')
                                reply_select(driver,'//*[@id="faultCause"]','其他原因')

                                try:
                                    #选择子告警手动清除时间为当前时间
                                    reply_subalert(driver)
                                except Exception:
                                    pass

                                #将driver切回主页面，下一步操作主页面上的元素
                                driver.switch_to.default_content()
                                
                                #提交并等待完成
                                reply_submit(driver)
                            
                            #回复接入网工单
                            if data_list_type[i-1] == '接入网':
                                j += 1

                                #开始回复工单
                                reply_order(driver,data_list_url,i,typeC,district)
                                reply_select(driver,'//*[@id="faultCause"]','设备温度超过门限')
                                
                                try:
                                    #选择子告警手动清除时间为当前时间
                                    reply_subalert(driver)
                                except Exception:
                                    pass

                                #将driver切回主页面，下一步操作主页面上的元素
                                driver.switch_to.default_content()
                                
                                #提交并等待完成
                                reply_submit(driver)

                            #回复互联网工单
                            if data_list_type[i-1] == '互联网':
                                j += 1

                                #开始回复工单
                                reply_order(driver,data_list_url,i,typeD,district)
                                reply_select(driver,'//*[@id="faultCause"]','供电系统故障')
                                
                                try:
                                    #选择子告警手动清除时间为当前时间
                                    reply_subalert(driver)
                                except Exception:
                                    pass

                                #将driver切回主页面，下一步操作主页面上的元素
                                driver.switch_to.default_content()
                                
                                #提交并等待完成
                                reply_submit(driver)

                            #回复动环网工单
                            if data_list_type[i-1] == '动环网':
                                j += 1

                                try:
                                    #开始回复工单
                                    reply_order(driver,data_list_url,i,typeE,district)
                                    reply_select(driver,'//*[@id="faultCause"]','模块故障')
                                except Exception:
                                    #回复结构为是否现场处理
                                    reply_select(driver,'//*[@id="isDisposal"]','否')      

                                try:
                                    #选择子告警手动清除时间为当前时间
                                    reply_subalert(driver)
                                except Exception:
                                    pass

                                #将driver切回主页面，下一步操作主页面上的元素
                                driver.switch_to.default_content()
                                
                                #提交并等待完成
                                reply_submit(driver)
                        except Exception as eA:
                            #关闭标签页
                            driver.close()
                            ur += 1
          
                            #切回当前标签页，使当前标签页作为可操作对象
                            driver.switch_to.window(driver.window_handles[-1])
                    

            #关闭浏览器窗口
            driver.quit()

            end_time = datetime.now()
            count += 1
            print(f"{datetime.now().replace(microsecond=0)}：{district}：第{count}次回单，用时{(end_time - start_time).seconds}秒,本次共检测到{len(data_list_url)}个工单,回复了{j}个工单,其中{ur}个回复失败")
            time.sleep(int(frequence)*60)
        except Exception as e1:
            print(f"{datetime.now().replace(microsecond=0)}：{district}：错误1：",e1)
            try:
                driver.quit()
            except Exception as e2:
                print(f"{datetime.now().replace(microsecond=0)}：{district}：关闭窗口错误：",e2)

#创建线程实例，方便在主函数中多线程创建
def t(u,p,f,d,m):
    a = threading.Thread(target=reply_gongdan,args=(u,p,f,d,m))
    a.daemon = True
    a.start()

def clear(f):
    while True:
        os.system('taskkill /F /IM firefox.exe')
        os.system('taskkill /F /IM geckodriver.exe')
        time.sleep(int(f)*60)

def c(f):
    a = threading.Thread(target=clear,args=(f,))
    a.daemon = True
    a.start()



