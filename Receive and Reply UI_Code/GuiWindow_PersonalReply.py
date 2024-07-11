import wx
from FunctionReply_Code import *
import sys
from pynput import mouse,keyboard

#wx.Frame是父类，定义子类CreatWindow
class CreatWindow(wx.Frame):
    def __init__(self,parent,title):
        wx.Frame.__init__(self,parent,title=title,size=(900,443))

        #窗口居中 (self指当前对象）
        self.Center()

        panel = wx.Panel(self)

        #添加容器，FlexGridSizer(rows,cols,vgap,hgap)
        #rows定义行数，cols定义列数，vgap定义垂直方向上行间距，hgap定义定义水平方向上列间距
        ls2 = wx.FlexGridSizer(2,3,20,5)
        ls3 = wx.BoxSizer(wx.HORIZONTAL)
        rs = wx.FlexGridSizer(2,1,3,10)
        rs.AddGrowableRow(1,1)
        rs.AddGrowableCol(0,1)
        
        
        #文字显示      
        t3=wx.StaticText(panel, label='日志信息：')
        t4=wx.StaticText(panel, label='回单频率：')
        t5=wx.StaticText(panel, label='分钟/次 ')
        t6=wx.StaticText(panel, label='后台清除频率：')
        t7=wx.StaticText(panel, label='分钟/次')                         
             
        minute1 = ['5','10','15','20','25','30']
        minute2 = ['120','180','240','300','360','420']

        #时间间隔下拉菜单设置
        self.Select_1 = wx.ComboBox(panel,-1,value = '10',choices = minute1,style = wx.CB_SORT,size=(40,25))

        #清除后台频率
        self.Select_2 = wx.TextCtrl(panel,-1,value = '240',size = (40,25))
        
        #日志信息文本框
        self.Log_Text = wx.TextCtrl(panel,style=wx.TE_MULTILINE|wx.TE_READONLY, size=(550,250))

        #开始按钮和停止按钮
        self.Button_Start = wx.Button(panel,-1,"开始")
        self.Button_Stop = wx.Button(panel,-1,"退出")  

        #border根据flag的值来设置边界(与其它控件距离）的大小，如wx.LEFT就是指左边边界;如果要控件自适应窗口大小，flag中一定要有wx.EXPAND
        #往容器添加控件,AddMany([(item1,proportion,flag,border),(item2,proportion,flag,border)...])
        ls2.AddMany([(t4,0,wx.EXPAND),(self.Select_1,0,wx.EXPAND),(t5,0,wx.EXPAND),
                     (t6,0,wx.EXPAND),(self.Select_2,0,wx.EXPAND),(t7,0,wx.EXPAND),])
        ls3.AddMany([(self.Button_Start,0,wx.EXPAND|wx.LEFT,35),(self.Button_Stop,0,wx.EXPAND|wx.LEFT,40)])
        rs.AddMany([(t3,0,wx.EXPAND|wx.TOP,10),(self.Log_Text,0,wx.EXPAND|wx.RIGHT|wx.BOTTOM,30)])

        #ls1,ls2，ls3垂直分布
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

        self.Select_2.Disable()


            
    def OnClose(self,event):
        r = wx.MessageDialog(self,"确定要关闭窗口?","警告",wx.YES_NO|wx.ICON_INFORMATION)

        #如果点击了“确定”
        if r.ShowModal() == wx.ID_YES:
            os.system('taskkill /F /IM geckodriver.exe')
            os.system('taskkill /F /IM firefox.exe')
            sys.exit()

              
    #点击开始按钮后的事件
    def Recieve_WO(self,event):
        try:
       
            #获取接单频率
            f = self.Select_1.GetValue()
            #获取后台清除频率
##            fb = self.Select_2.GetValue()
            driver = webdriver.Firefox()
            driver.get("http://10.93.19.175:8091/wyeoms/")            
            s = wx.MessageDialog(None,"请登陆至陕西联通综合调度系统首页后，再点击“是”开始程序运行","提示",wx.YES_NO|wx.ICON_INFORMATION)                
            if s.ShowModal() == wx.ID_YES:                
                t(driver,f)               
                #禁用开始按钮、下拉菜单
                self.Button_Start.Disable()
                self.Select_1.Disable()            

            else:
                os.system('taskkill /F /IM firefox.exe')
                os.system('taskkill /F /IM geckodriver.exe')   

            
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
        os.system('taskkill /F /IM firefox.exe')
        os.system('taskkill /F /IM geckodriver.exe')
        sys.exit()
        
        
if __name__ == '__main__':
    app = wx.App(False)
    MyWindow = CreatWindow(None,"亿阳系统网络工单回复")

    #通过show（）方法激活框架窗口
    MyWindow.Show(True)

    #输入Application对象的主事件循环
    #1、mainloop()方法允许程序循环执行，并进入等待和处理事件。
    #窗口中的组件可以理解为一个连环画
    #2、mainloop()方法的作用是监控每个组件，当组件发生变化或触发事件时，会立即更新窗口。
    app.MainLoop()


