from datetime import datetime
import re
from selenium.webdriver.common.by import By
import requests
from lxml import etree

#获取网页html格式源码
def get_html_str(pagesource):
    """
    定义一个函数, 新建一个空变量html_str， 请求网页获取网页源码，如果请求成功，则返回结果，如果失败则返回空值
    url: 入参参数, 指的是我们普通浏览器中的访问网址
    """
    html_str = ""
    try:
        """使用三方库etree把网页源码字符串转换成HTML格式"""
        html_str = etree.HTML(pagesource)
    except Exception as e:
        print(e)
    return html_str


#获取工单页面第3列“工单流水号”的数据
def get_url(html_str):
    """
    定义一个函数, 新建一个变量pdata_list初始值为空列表（也可以叫空数组）， 在网页源码中匹配出每一行的内容
    html_str: 入参参数, 指的是网页源码，HTML格式的
    """
    data_list = []
    
    try:
        option = html_str.xpath('/html/body/div[4]/div/div/table/tbody/tr')
        for op in option:

            """获取第3列（工单流水号）的链接"""
            try:
                col3 =op.xpath("./td[3]/a/@href")[0]
                
            except Exception as error:
                print(error)
                col3=""

            data_list.append(col3)                            
    except Exception as e:
        print(e)
    return data_list


#获取工单页面第6列“派单时间”的数据
def get_time(html_str): 
    data_list = []
    try:
        option = html_str.xpath('/html/body/div[4]/div/div/table/tbody/tr')
        for op in option:
            try:
                #将取到的时间字符串改为时间格式
                col8 = datetime.strptime(re.findall('\d{4}-\d{1,25}-\d{1,2}\s+\d{1,2}:\d{1,2}:\d{1,2}',op.xpath("./td[6]/text()")[0])[0],'%Y-%m-%d %H:%M:%S')
            except Exception:
                col8 = ""
            data_list.append(col8)
    except Exception as e:
        print(e)
    return data_list
    

#获取工单页面第8列“专业”的数据
def get_type(html_str):
    data_list = []
    try:
        option = html_str.xpath('/html/body/div[4]/div/div/table/tbody/tr')
        for op in option:
            try:
                col8 = re.findall('[\\u4e00-\\u9fa5]+',op.xpath("./td[8]/text()")[0])[0]               
            except Exception as error:
                print(error)
                col8 = ""
            data_list.append(col8)
    except Exception as e:
        print(e)
    return data_list

#获取接工单页面第11列“任务所有者”数据
def get_data(html_str):
    """
    定义一个函数, 新建一个变量data_list初始值为空列表（也可以叫空数组）， 在网页源码中匹配出每一行的内容
    html_str: 入参参数, 指的是网页源码，HTML格式的
    """
    data_list = []
    
    try:
        option = html_str.xpath('/html/body/div[4]/div/div/table/tbody/tr')
        for op in option:

            """获取第11列（任务所有者）的值"""
            try:
                col11 = re.findall('[\\u4e00-\\u9fa5]+',op.xpath("./td[11]/text()")[0])[0]
            except:
                col11=""
            data_list.append(col11)                            
    except Exception as e:
        print(e)
    return data_list

#获取工单页面第13列“上一级处理人员”数据
def get_personnel(html_str):
    data_list = []
    try:
        option = html_str.xpath('/html/body/div[4]/div/div/table/tbody/tr')
        for op in option:        
            try:
                col14 = re.findall('[\\u4e00-\\u9fa5]+',op.xpath("./td[13]/text()")[0])[0]
            except Exception:
                col14=""
            data_list.append(col14)
    except Exception as e:
        print(e)
    return data_list

#获取工单页面第14列"告警清除时间”数据
def get_alarm(html_str):
    data_list = []
    try:
        option = html_str.xpath('/html/body/div[4]/div/div/table/tbody/tr')
        for op in option:        
            try:
                col14 = re.findall('[\\u4e00-\\u9fa5]+',op.xpath("./td[14]/text()")[0])[0]
            except Exception:
                col14=""
            data_list.append(col14)
    except Exception as e:
        print(e)
    return data_list

#获取工单总页数
def get_pagenum(html_str):
    try:
        #获取最后一页的页数
        option = html_str.xpath('/html/body/div[4]/div/div/span[2]/strong')
        num = re.findall('\d+',option[0].xpath("./text()")[0])[0]
    except Exception:
        num = ""
        print("获取页数错误")
        
    return num
    

#获取网络故障3.0待办工单页面工单列表中第8列数据
def get_processing_stage(html_str):
    data_list = []
    try:
        option = html_str.xpath('//*[@id="app"]/section/section/main/div[1]/div[2]/div/div/div[2]/div/div[3]/table/tbody/tr')

        for op in option:
            try:
                col8 = re.findall('[\\u4e00-\\u9fa5]+',op.xpath("./td[8]/div/text()")[0])[0]               

            except Exception as e:        
                print(e)
                col8 = ""
                
            data_list.append(col8)
            
    except Exception as e:
        print(e)
        
    return data_list
            
            






