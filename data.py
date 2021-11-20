import xlrd
from datetime import date, datetime

# 指定起始日,可用'date.today()'代替當天(把下一行的'#'拿掉即可)
# targetDate = date.today()
targetDate = datetime.strptime('20041231',"%Y%m%d")
# 指定日期五年後
fiveYear =  targetDate.replace(year = targetDate.year + 5)

# 指定大於數
biggerThan = 7

# 匯入excel檔
data = xlrd.open_workbook('現金股利率.xlsx')
table = data.sheets()[0]

# 建立匯出目的地
path = 'result.txt'
# 開啟檔案
f = open(path, 'w')

for j in range(1,table.ncols):
    # 總和
    sum = 0
    # 次數
    count = 0
    # 平均數
    average = 0
    for i in range(1,table.nrows):
        # 取出利率和當天日期
        listDate = datetime.strptime(str(int(table.cell(i,0).value)),"%Y%m%d")
        rate = table.cell(i,j).value

        # 以0替代欄位空值
        if rate=='' : rate=0

        if(listDate<fiveYear and (isinstance(rate,float) or isinstance(rate,int))):
            # 若符合日期條件, 增加總和
            sum += rate
            # 次數+1
            count += 1

        # 計算平均數
        if(count!=0):
            average = sum/count
        
    if(average>biggerThan):
        # 若平均數符合條件(大於7)
        message = '股票代號:' + str(int(table.cell(0,j).value)) + ' 五年平均大於'+ str(biggerThan) + ', 股利為:' + str(average) + '\n'    
        # 顯示在終端機
        print(message)
        # 寫入檔案
        f.write(message)

# 關閉檔案
f.close()