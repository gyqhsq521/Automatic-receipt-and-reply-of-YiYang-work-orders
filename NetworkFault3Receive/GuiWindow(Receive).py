
import wx
from NetworkFault3receive import *
import threading
import sys
import os
import time

#wx.Frame是父类，定义子类CreatWindow
class CreatWindow(wx.Frame):
    def __init__(self,parent,title):
        wx.Frame.__init__(self,parent,title=title,size=(900,443))

        #窗口居中 (self指当前对象）
        self.Center()

        panel = wx.Panel(self)

        #添加容器，FlexGridSizer(rows,cols,vgap,hgap)
        #rows定义行数，cols定义列数，vgap定义垂直方向上行间距，hgap定义定义水平方向上列间距
        ls2 = wx.FlexGridSizer(3,3,20,5)
        ls3 = wx.BoxSizer(wx.HORIZONTAL)
        rs = wx.FlexGridSizer(2,1,3,10)
        rs.AddGrowableRow(1,1)
        rs.AddGrowableCol(0,1)
        
        
        #文字显示
        t3=wx.StaticText(panel, label='日志信息：')
        t4=wx.StaticText(panel, label='接单频率：')
        t5=wx.StaticText(panel, label='分钟/次 ')
        t6=wx.StaticText(panel, label='新创建端口号')
        t7=wx.StaticText(panel, label='（注：未被占用）')
        t8=wx.StaticText(panel, label='无浏览器窗口：')
           
             
        minute1 = ['5','10','15','20','25','30']
        minute2 = ['120','180','240','300','360','420']
        choice = ["是","否"]

        #时间间隔下拉菜单设置
        self.Select_1 = wx.ComboBox(panel,-1,value = '7',choices = minute1,style = wx.CB_SORT,size=(40,25))

        #新建窗口端口号设置
        self.Select_2 = wx.TextCtrl(panel,-1,value = '9222',size = (40,25))

        #无窗口模式选择下拉菜单设置
        self.Select_3 = wx.ComboBox(panel,-1,value = '否',choices = choice,style = wx.CB_SORT,size = (40,25))
        
        #日志信息文本框
        self.Log_Text = wx.TextCtrl(panel,style=wx.TE_MULTILINE|wx.TE_READONLY, size=(550,250))

        #开始按钮和停止按钮
        self.Button_Start = wx.Button(panel,-1,"开始")
        self.Button_Stop = wx.Button(panel,-1,"退出")  

        #border根据flag的值来设置边界(与其它控件距离）的大小，如wx.LEFT就是指左边边界;如果要控件自适应窗口大小，flag中一定要有wx.EXPAND
        #往容器添加控件,AddMany([(item1,proportion,flag,border),(item2,proportion,flag,border)...])
        ls2.AddMany([(t4,0,wx.EXPAND),(self.Select_1,0,wx.EXPAND),(t5,0,wx.EXPAND),
                     (t6,0,wx.EXPAND),(self.Select_2,0,wx.EXPAND),(t7,0,wx.EXPAND),
                     (t8,0,wx.EXPAND),(self.Select_3,0,wx.EXPAND)])
        ls3.AddMany([(self.Button_Start,0,wx.EXPAND|wx.LEFT,35),(self.Button_Stop,0,wx.EXPAND|wx.LEFT,40)])
        rs.AddMany([(t3,0,wx.EXPAND|wx.TOP,10),(self.Log_Text,0,wx.EXPAND|wx.RIGHT|wx.BOTTOM,30)])

        #ls2，ls3垂直分布
        ls = wx.FlexGridSizer(2,1,20,0)
        ls.AddMany([(ls2,0,wx.EXPAND|wx.LEFT|wx.TOP,40),(ls3,0,wx.EXPAND|wx.LEFT|wx.TOP,20)])

        #ls,rs水平分布
        sa = wx.FlexGridSizer(1,2,10,10)
        sa.AddMany([(ls,0,wx.EXPAND|wx.TOP,20),(rs,0,wx.EXPAND|wx.LEFT,35)])

        #第一行第二列能够自适应窗口大小，参数（索引，proportion）
        sa.AddGrowableRow(0, 1)
        sa.AddGrowableCol(1, 1)

        #启用sa布局
        panel.SetSizer(sa)
        

        #将执行功能的函数和按钮绑定
        self.Bind(wx.EVT_BUTTON,self.Recieve_WO, self.Button_Start)
        self.Bind(wx.EVT_BUTTON,self.Stop_Recieve_WO,self.Button_Stop)

        #主窗口绑定自身关闭事件(点击右上角的叉后，执行OnClose函数）
        self.Bind(wx.EVT_CLOSE,self.OnClose)

        #把在shell中输出的信息重定向到日志信息文本框
        sys.stdout=self.Log_Text

        self.Select_3.Disable()

    #光标移动防止锁屏
    def MouseMov(self):
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

    
    def OnClose(self,event):
        r = wx.MessageDialog(self,"确定要关闭窗口?","警告",wx.YES_NO|wx.ICON_INFORMATION)

        #如果点击了“确定”
        if r.ShowModal() == wx.ID_YES:
            os.system('taskkill /F /IM chromedriver.exe')   
            sys.exit()

              
    #点击开始按钮后的事件
    def Recieve_WO(self,event):
        try:
            #获取接单频率
            frequence = self.Select_1.GetValue()

            #获取后台清除频率
            pid = self.Select_2.GetValue()
            
            #选择是否无窗口运行
            mode_type = self.Select_3.GetValue()
            
            #清除后台的线程
            os.system('taskkill /F /IM chromedriver.exe')
            os.system('taskkill /F /IM chrome.exe')

            #打开一个远程调试模式的chrome
            os.popen(f'start chrome --remote-debugging-port={pid} --user-data-\dir="D:\selenium" http://sso.portal.unicom.local/eip_sso/aiportalLogin.html?appid=na186&oawx_t=A0002')


            s = wx.MessageDialog(self,"请登陆至指定页面后，再点击“是”开始程序运行","警告",wx.YES_NO|wx.ICON_INFORMATION)
            #如果点击了“确定”
            if s.ShowModal() == wx.ID_YES:
                #防锁屏
                self.a = threading.Thread(target=self.MouseMov)
                self.a.daemon = True
                self.a.start()
     
                #防止函数Recieve_gongdan执行时，界面无响应
                #给Recieve_gongdan函数新增一个线程，该函数和GUI各自执行。args元组里面写该函数的参数
                self.b = threading.Thread(target=Receive,args=(frequence,pid,mode_type))
                #如果某个子线程的daemon属性为False，主线程结束时会检测该子线程是否结束，如果该子线程还在运行，则主线程会等待它完成后再退出；
                #如果某个子线程的daemon属性为True，主线程运行结束时不对这个子线程进行检查而直接退出，同时所有daemon值为True的子线程将随主线程一起结束，而不论是否运行完成。
                self.b.daemon = True
                self.b.start()

                #禁用开始按钮、下拉菜单
                self.Button_Start.Disable()
                self.Select_1.Disable()
                self.Select_2.Disable()
                
                
            else:
                os.system('taskkill /F /IM chrome.exe')          
    
        except Exception as e:

            print(e)

            # 语法是(self, 内容, 标题, ID)
            dlg = wx.MessageDialog(self, "{}".format(e), "错误", wx.OK)

            # 显示对话框
            dlg.ShowModal()

            # 当结束之后关闭对话框
            dlg.Destroy()  

    #停止接收工单
    def Stop_Recieve_WO(self,event):

        #清除后台残留进程
        os.system('taskkill /F /IM chrome.exe')
        os.system('taskkill /F /IM chromedriver.exe')
        sys.exit()


        
        

        
if __name__ == '__main__':

    app = wx.App(False)
    MyWindow = CreatWindow(None,"网路故障3.0工单接收")

    #通过show（）方法激活框架窗口
    MyWindow.Show(True)

    #输入Application对象的主事件循环
    #1、mainloop()方法允许程序循环执行，并进入等待和处理事件。
    #窗口中的组件可以理解为一个连环画
    #2、mainloop()方法的作用是监控每个组件，当组件发生变化或触发事件时，会立即更新窗口。
    app.MainLoop()


