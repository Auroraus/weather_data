# -*- coding: utf-8 -*-
"""
Created on Tue Nov 14 22:42:43 2017

@author: zf
"""

import time
from tkinter import messagebox
import base64,os,json
from panda import img
import win32com.client
import requests
import tkinter
import xlwt


w=xlwt.Workbook()
h=20
sheet =w.add_sheet("气象情况")
style = xlwt.easyxf('font: bold 1, color red;')
sheet.row(0).height = 256 *3
sheet.write(0,0,"日期",style)
sheet.write(0,0+1,"天气状况",style)
sheet.write(0,1+1,"温度【单位：摄氏度】",style)
sheet.write(0,2+1,"表观温度【单位：摄氏度】",style)
sheet.write(0,3+1,"露点【单位：摄氏度】",style)
sheet.write(0,4+1,"湿度",style)
sheet.write(0,5+1,"大气压【单位：千帕】",style)
sheet.write(0,6+1,"风速【单位：m/s】",style)
sheet.col(0).width = 256 *h
sheet.col(1).width = 256 *h
sheet.col(2).width = 256 *h
sheet.col(3).width = 256 *h
sheet.col(4).width = 256 *h
sheet.col(5).width = 256 *h
sheet.col(6).width = 256 *h
sheet.col(7).width = 256 *h

r=tkinter.Tk() #tkinter root初始化 
r.title('气象数据获取')#界面标题栏

tmp = open("tmp.ico","wb+")
tmp.write(base64.b64decode(img))
tmp.close()
r.iconbitmap("tmp.ico")
os.remove("tmp.ico")
r.geometry()#界面大小（自适应）

mu=tkinter.Menu(r)
fi=tkinter.Menu(mu,tearoff=False)
fi.add_command(label='开发者：M-45',command='callback')
fi.add_command(label='版本号：1.1.0 ',command='callback')
fi.add_command(label='软件介绍：用来帮老杜获取气象数据',command='callback')
fi.add_command(label='退出',command=r.destroy)
mu.add_cascade(label='关于',menu=fi)
r.config(menu=mu)


tkinter.Label(r,text='谨以此程序献给我的好友----杜永峰，祝你使用愉快！').pack()#第一个标签

tkinter.Label(r,text='KEY[请不要改动]').pack()#第一个标签
input110=tkinter.StringVar()#捕获用户输入
xen110=tkinter.Entry(r,textvariable=input110,width=25)#用户文本输入
input110.set('32074fb694acf1bb183c4d6f07c13352')#输入框预设值
xen110.pack()#使用户输入框生效

def say():
    sayword='您好，欢迎您使用本程序，利用本程序您可以获取你想要的气象数据。注意输入目标地的经纬度【精确至小数点后三位】，祝您使用愉快，若有技术问题，请发到我的邮箱：zf083415@gmail.com'
    s = win32com.client.Dispatch("SAPI.SpVoice")
    s.Speak(sayword)
    time.sleep(1)

tkinter.Button(r,text=("初次使用前请点击"),command=say,width=15,height=1,bg='green').pack()        

    


tkinter.Label(r,text='请输入你要爬取的城市经度').pack()#第一个标签
input21=tkinter.StringVar()#捕获用户输入
xen21=tkinter.Entry(r,textvariable=input21,width=25)#用户文本输入
input21.set('117')#输入框预设值
xen21.pack()#使用户输入框生效
             
tkinter.Label(r,text='请输入你要爬取的城市维度').pack()#第一个标签
input1=tkinter.StringVar()#捕获用户输入
xen=tkinter.Entry(r,textvariable=input1,width=25)#用户文本输入
input1.set('30')#输入框预设值
xen.pack()#使用户输入框生效

tkinter.Label(r,text='请输入您要查找的数据的起始年份').pack()#第一个标签
input2=tkinter.StringVar()#捕获用户输入
xen1=tkinter.Entry(r,textvariable=input2,width=25)#用户文本输入
input2.set('2015')#输入框预设值
xen1.pack()#使用户输入框生效

tkinter.Label(r,text='请输入您要查找的数据的截止年份').pack()#第一个标签
input8=tkinter.StringVar()#捕获用户输入
xen8=tkinter.Entry(r,textvariable=input8,width=25)#用户文本输入
input8.set('2016')#输入框预设值
xen8.pack()#使用户输入框生效

tkinter.Label(r,text='请输入您要查找的数据的起始月份').pack()#第一个标签
input9=tkinter.StringVar()#捕获用户输入
xen9=tkinter.Entry(r,textvariable=input9,width=25)#用户文本输入
input9.set('1')#输入框预设值
xen9.pack()#使用户输入框生效
        
tkinter.Label(r,text='请输入您要查找的数据的截止月份').pack()#第一个标签
input10=tkinter.StringVar()#捕获用户输入
xen10=tkinter.Entry(r,textvariable=input10,width=25)#用户文本输入
input10.set('2')#输入框预设值
xen10.pack()#使用户输入框生效



def start():
  n=1
  key=input110.get()
  longitude=input21.get()
  latitude=input1.get()
  sy=int(input2.get())
  ey=int(input8.get())
  sm=int(input9.get())
  em=int(input10.get())
  messagebox.showinfo('提示：','即将开始爬取【一天数据大约需要半秒，时间自己算】')
  weizhi=latitude+'-'+longitude+str(sy)+'.'+str(sm)+'.'+str(ey-1)+'.'+str(em-1)+'_气象数据.xls'
  for year in range(sy,ey):
    for month in range(sm,em):
      for day in range(1,32):
        date=str(year)+'-'+str(month)+'-'+str(day)+' '+'14:00'
        tim=int(time.mktime(time.strptime(date,'%Y-%m-%d %H:%M')))
        try:
#https://api.darksky.net/forecast/32074fb694acf1bb183c4d6f07c13352/30.000,117.0,1510556488?lang=zh&units=si&exclude=daily
            url='https://api.darksky.net/forecast/'+key+'/'+str(latitude)+','+str(longitude)+','+str(tim)+'?lang=zh&units=si&exclude=daily'
            head={"user-agent":"Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 UBrowser/6.2.3637.220 Safari/537.36"}
            r=requests.get(url,headers=head)
            r.encoding=r.apparent_encoding
            data=json.loads(r.text)
            k=(data['currently'])
            if 1:
                        sheet.write(n,0,time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(k['time'])))
                        sheet.write(n,1,(k['summary']))
                        sheet.write(n,2,(k['temperature']))
                        sheet.write(n,3,(k['apparentTemperature']))
                        sheet.write(n,4,k['dewPoint'])
                        sheet.write(n,5,k['humidity'])
                        sheet.write(n,6,k['pressure'])
                        sheet.write(n,7,k['windSpeed'])
                        n=n+1
        except:
            pass
  
  if(n>2):
      w.save(weizhi)
      messagebox.showinfo('提示','：数据爬取完成，在'+weizhi)
tkinter.Button(r,text=("开始爬取"),command=start,width=10,height=1).pack()      
r.mainloop()      
