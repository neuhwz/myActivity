# 由TaskList中的任务及自己的活动经历得到某日（今日）的活动，并写入Activity
import xlwings as xw       # 导入操作Excel的第三方模块xlwings
import datetime            # 导入获取今日日期的内置模块
import pandas as pd        # 导入数据分析模块
thisDate=datetime.date.today()   # 获取今天日期
# thisDate=datetime.date.today()-datetime.timedelta(days=1) #前几天的日期，今天、昨天、前天分别取days=0、1、2
print(thisDate)

app=xw.App()               # 初始化
app.display_alerts=False   # 不显示Excel消息框
app.screen_updating=False  # 关闭屏幕更新,可加快宏的执行速度
wb=app.books.open('myLife.xlsx') # 打开人生导航的Excel工作簿文件

sht = wb.sheets["Activity"]    # 实例化一个工作表对象
NewRow = sht.used_range.last_cell.row + 1  # 将从Activity最后填写该日的活动
print(NewRow)

iCnt=0 
taskSht=wb.sheets("TaskList")  # 实例化TaskList工作表对象
lastRow = taskSht.used_range.last_cell.row + 1
print("TaskList",lastRow)
# 扫描TaskList表，看是否有该日的任务。如果有则将当日任务加到活动中
for j in range(2,lastRow):      # 在整个TaskList数据中找到该日期数据

    if taskSht.range(j, 1).value.strftime("%m/%d/%Y") == thisDate.strftime("%m/%d/%Y"): #转换成统一格式比较
        sht.range(NewRow, 1).value = taskSht.range(j, 1).value #日期
        sht.range(NewRow, 2).value = taskSht.range(j, 2).value #开始时间
        sht.range(NewRow, 3).value = taskSht.range(j, 3).value #操作
        sht.range(NewRow, 4).value = taskSht.range(j, 4).value #时长
        sht.range(NewRow, 5).value = taskSht.range(j, 5).value #对象名
        sht.range(NewRow, 7).value = taskSht.range(j, 6).value #说明
        NewRow+=1
        iCnt+=1
        print("上课",iCnt)

if NewRow<300:        #由MyAction工作表生成当日活动
    actionSht=wb.sheets("MyAction") # 实例化MyAction工作表对象
    for j in range(2,32-iCnt):      
        sht.range(NewRow, 1).value = thisDate #日期
        sht.range(NewRow, 3).value = actionSht.range(j, 1).value #操作
        sht.range(NewRow, 6).value = actionSht.range(j, 3).value #数量
        sht.range(NewRow, 4).value = (int(actionSht.range(j, 4).value)// 5) * 5 #时长
        NewRow+=1
else:                 #由Activity中的过去活动经历，生成当日活动
    df = pd.read_excel('myLife.xlsx',sheet_name = 'Activity') # 获取Activity表所有数据
    df = df.groupby('操作',as_index=False).agg({'日期':'count','时长':'mean', '数量':'mean'})
    print(df)
    df.drop(df[df['操作']=='上课'].index,inplace=True)
    print(df)
    df.sort_values(['日期'],ascending=False, inplace=True)
    for j in range(30-iCnt):      
        sht.range(NewRow, 1).value = thisDate #日期
        sht.range(NewRow, 3).value =  df.iat[j,0]#操作
        sht.range(NewRow, 6).value =  df.iat[j,3]#数量
        sht.range(NewRow, 4).value =  df.iat[j,2]#时长
        NewRow+=1

wb.save()   # 保存Excel文件
wb.close()  # 关闭Excel文件
app.quit()  # 退出excel程序