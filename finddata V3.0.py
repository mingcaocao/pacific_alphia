import requests
import urllib2
from bs4 import BeautifulSoup
import os
import xlwt




def openfile(name):
    newname=''
    index=name.index(' ')
    newname+=(name[:index:])+' '
    name=name[index+1::]
    index=name.index(' ')
    newname+=(name[:index:])
    os.makedirs('~/Document/stock-data/%s'%newname)

def writexls(data,number,name):
    newname=''
    index=name.index(' ')
    newname+=(name[:index:])+' '
    name=name[index+1::]
    index=name.index(' ')
    newname+=(name[:index:])
    year=2017+(number-1)//4
    number=number%4
    if number==0:
        number+=4
    print(number,year)
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('sheet 1')
    for n in range(len(data)):
        for m in range(len(data[n])):
            if m==3 or m==4:
                string=''
                for j in range(len(data[n][m])):
                    try:
                        int(data[n][m][j])
                        string+=data[n][m][j]
                    except:continue
                try:sheet.write(n,m,int(string))
                except:
                    sheet.write(n,m,data[n][m])
                    print('!')
            else:sheet.write(n,m,data[n][m])

    workbook.save('~/Document/stock-data/%s/%d%d.xls'%(newname,year,number))




def getdata(numlist):
    for num in numlist:
        payload = {'action':'getcompany','filenum':'028-'+num}
        r = requests.get("https://www.sec.gov/cgi-bin/browse-edgar", params=payload)
        page = urllib2.urlopen(r.url)
        soup = BeautifulSoup(page)
        html = soup.find_all('a')
        number=2
        label=1
        for h in html[12:-3:]:
            dochtml = h.get('href')
            midpage = urllib2.urlopen('https://www.sec.gov'+dochtml)
            midsoup=BeautifulSoup(midpage)
            midhtml = midsoup.find_all('a')
            web=midhtml[10]
            getweb=web.get('href')
            namehtml = midsoup.find(class_='companyName')
            name=namehtml.text
            objpage = urllib2.urlopen('https://www.sec.gov'+getweb)
            soup2=BeautifulSoup(objpage)
            tbody = soup2.find('tbody')
            i=1
            n=0
            lst=[]
            numlist=[1,2,3,4,5,7]
            while tbody != None:
                lst.append([])
                for j in numlist:
                    try:
                        lst[n].append(str((tbody.contents[i]).contents[j].text))
                        if str((tbody.contents[i]).contents[j].text) in ['Call','Put','COLUMN 7','CALL']:
                            del lst[n]
                    except:
                        n+=1
                        continue
                i+=2
                try:tbody.contents[i]
                except:
                    print('done')
                    break
            if label == 1 :
                openfile(name)
            writexls(lst,number,name)
            number-=1
            label+=1
            if number==-13:
                break

        print('finish')

if __name__=='__main__':
    numlist=['16256','10684']
    numlist1=['14843','14564','05497','05369','14996','03499','07120','02408','11152','06270','15008','11694','11096','16256','10684','07484','10134']
    getdata(numlist)
