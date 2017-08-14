# -*- coding: utf-8 -*-
"""
Created on Fri Aug 11 09:31:15 2017
@author: apple
"""
from __future__ import division
import matplotlib.pyplot as plt
import xlrd
from datetime import datetime
import matplotlib.dates as mdates


def readcompany():
    companydata=[]
    book = xlrd.open_workbook(r'C:\Users\apple\Desktop\stockdata.xls')
    count = len(book.sheets())
    for i in range(count):
        companydata.append([])
        sheet=book.sheets()[i]
        companyname=sheet.name
        companydata[i].append([companyname])
        nrows = sheet.nrows    #行总数
        for j in range(nrows-1,0,-1):
            companydata[i].append([sheet.cell_value(j,0),sheet.cell_value(j,2)])#key为日期 value为stock price
    return(companydata)

def readfund():
    funddata=[]
    book = xlrd.open_workbook(r'C:\Users\apple\Desktop\funddata.xlsx')     #第一列是股票代码 第二列是购入时间第三列是difference
    sheet=book.sheets()[0]
    nrows = sheet.nrows    #行总数
    for i in range(nrows):
        funddata.append([sheet.cell_value(i,0),transfertime(sheet.cell_value(i,1)),sheet.cell_value(i,2)])#输入一组股票代码,购入时间和difference
    return(funddata)

def draw(companydata,funddata):
    #companydata[company[companyname,{},{},{}],[]...]
    for i in range(len(companydata)):
        x=[]
        y=[]
        for n in range(1,len(companydata[i])):
            x.append(companydata[i][n][0])
            y.append(companydata[i][n][1])
        x=transferdate(x)
        topy=top(y)
        downy=down(y)
        xs = [datetime.strptime(d, '%Y-%m-%d').date() for d in x]
        plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
        plt.gca().xaxis.set_major_locator(mdates.DayLocator())
        title=companydata[i][0]
        plt.figure(figsize=(25,15))
        plt.plot(xs,y,linewidth=1)
        plt.xlabel('date')
        plt.ylabel('stock price')
        plt.ylim(downy/1.05,topy*1.05)
        timelist=['20130815','20131115','20140215','20140515','20140815','20141115','20150215','20150515','20150815','20151115','20160215','20160515','20160815','20161115','20170215','20170515']
        #
        title=str((companydata[i][0]))[3:-2:]
        for j in range(len(funddata)):
            if title == funddata[j][0]:
                if int(funddata[j][2]) < 0:c='red'
                else:c='green'
                length=int(funddata[j][2])
                topline=1/4*abs(length)
                if topline>1:topline=1
                plt.axvline(funddata[j][1],ymin=0,ymax=topline,label='%s'%str(length),color=c,linestyle="--")
                for timelabel in timelist:
                    plt.axvline(timelabel,ymin=0,ymax=0.1,color='black',linestyle="--")
        plt.title('TS stock price of %s'%title)
        foo_fig = plt.gcf()
        foo_fig.savefig(r'C:\Users\apple\Desktop\graphs_test\%s.png'%title, dpi=1000)
        plt.legend()
        plt.show()
        print(title)


def top(y):
    maxy=0
    #如果读出来的数据是字符串
    for i in range(len(y)):
        if float(y[i]) > maxy:
            maxy= float(y[i])
    return(maxy)


def down(y):
    try:
        miny=float(y[0])
    #如果读出来的数据是字符串
        for i in range(len(y)):
            if float(y[i]) < miny:
                miny= float(y[i])
        return(miny)
    except:return(0)

def transfertime(cell):
    timedic={'Q1':'0215','Q2':'0515','Q3':'0815','Q4':'1115'}
    part1=cell[:4:]
    part2=timedic[(cell[4::])]
    newcell=part1+part2
    return(newcell)

def transferdate(date):#将May 03,2017转成20170503
   newdate=[]
   monthdic={'Jan':'01','Feb':'02','Mar':'03','Apr':'04','May':'05','Jun':'06','Jul':'07','Aug':'08','Sep':'09','Oct':'10','Nov':'11','Dec':'12'}
   for i in range(len(date)):
       part1=(date[i])[:3:]
       part1=monthdic[part1]
       part2=(date[i])[4:6:]
       part3=(date[i])[-4::]
       newdate.append(part3+'-'+part1+'-'+part2)
   return(newdate)


if __name__=="__main__":
    companydata=readcompany()
    funddata=readfund()
    draw(companydata,funddata)
