import urllib, urllib.request
import xlrd, re,time,os,winsound
import pandas as pd
from playsound import playsound


osss = os.getcwd()
bbb = os.listdir(osss)
for i in range(len(bbb)):
  c = re.search('xls',bbb[i])
  if c:
    ccc = osss+'/'+bbb[i]

print(ccc)
#excel=xlrd.open_workbook('c:/20180730强平打印表.xls')
excel = xlrd.open_workbook(ccc)
data = excel.sheet_by_name("结算价")



#print (time1)
##########################################################################################################################
def get_stock_info(stock_no, num_retries=2):
    # try:
    # url = 'http://nufm.dfcfw.com/EM_Finance2014NumericApplication/JS.aspx?type=CT&cmd='+stock_no.strip()+'2&sty=CTBF&st=z&sr=&p=&ps=&cb=var%20pie_data=&js=(x)&token=28758b27a75f62dc3065b81f7facb365&_=1496312544427'
    url = 'http://hq.sinajs.cn/?_=1531364348612/&list=nf_' + stock_no.strip()
    #url = "https://hq.sinajs.cn/rn=" + str(round(time.time() * 1000)) + "&list=" + stock_no.strip()
    headers = {'User-agent': 'WSWP', 'referer': 'http://finance.sina.com.cn'}
    request = urllib.request.Request(url, headers=headers)#形成一个对象
    page = urllib.request.urlopen(request)
    page_content = page.read().decode('gbk')
    return page_content
##########################################################################################################################
b=[]
a=list(set(zip(data.col_values(4),data.col_values(7),data.col_values(8))))#提取4=合约号，7=90%涨幅，8=90%跌幅
#print(a)

for i in range(len(a)):#删除多余的内容，保留需要监控的合约
   a[i] = list(a[i])
   ####找出不要的期权合约，其它多余内容
   m = re.match(r'([a-z,A-Z]+)',a[i][0])
   if  not m:
      b.append(a[i])
   n = re.search('-', a[i][0])
   if n:
       b.append(a[i])
   ####将合约字母变为大写
   d = re.findall(r'([a-z]+)([0-9][0-9]+)', a[i][0])
   if d:
       a[i][0] = d[0][0].upper() + d[0][1]
for j in range(len(b)):###删除多余内容
        a.remove(b[j])

print(len(a))
#data_hq_1=get_stock_info('RB1907')
#print(data_hq_1.split(',')[8])
#########################################################################################################################
cc=[]
dd=[]
for i in range(len(a)):
    cc.append(0)
    dd.append(0)
cc_counter=0

while 1>0:
   time1 = str(time.strftime('%H:%M:%S', time.localtime(time.time())))
   
   for j in range(len(a)):
       data_hq = get_stock_info(a[j][0])
       cu = data_hq.split(',')
       if len(cu)<=1:  
           if dd[j]<1:
              #for x in range(3):
              print('           '+a[j][0]+': 脱离监控，请单独设置')
              dd[j]+=1

           continue
       data_hq_1=float(cu[8])
       #print(cu[0],'/',cu[8],a[j][1])
       if data_hq_1 >= a[j][1]:
          print( time1,'  '+a[j][0]+'   '+'上涨 超过90%')
          cc[j]=cc[j] + 1
       elif  data_hq_1 <= a[j][2]:
           print( time1,'  '+a[j][0]+'   '+'下跌 超过90%')
           cc[j] = cc[j] + 1
        #playsound('C:/untitled1/10367.wav')
       else:
          cc[j]=0
          continue
       cc_counter=cc[j]+cc_counter
   if cc_counter:
        print(' ')
        winsound.PlaySound('C:/untitled1/10367.wav',flags=0)#flags=0,winsound.SND_ALIAS  32位电脑上用后面的参数
