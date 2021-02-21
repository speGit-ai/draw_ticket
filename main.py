import sys
import os

import filepathread
from Classifi import class_text
import xlwings as xw

draw = []
ticket = []
fileNamePath = []
exist = 0
if __name__ == '__main__':

    excel = xw.App(visible=False, add_book=False) #表格可见性
    excel.display_alerts = False
    excel.screen_updating = False   #表格即时刷新

    LocalPath = os.getcwd()         #获取当前py文件夹
    filepathread.listdir(LocalPath, fileNamePath)   #获取当前文件夹所有文件列表

    #查找含明信的excel表
    for filename in fileNamePath:

        if (filename.endswith('xls') or filename.endswith('xlsx')) and ('明信水印长滩' in filename):
            filepath = os.path.join(LocalPath, filename)
            exist = 1
            break
    if exist == 0:
        print("没找到excel文件")
        sys.exit(0)



    #filepath = r'C:\Users\budget\Desktop\明信水印长滩12.25-1.25.xlsx'
    try:
        workbook = xw.books.active #绑定已打开文件
    except:
        workbook = excel.books.open(filepath)   #打开文件 路径：filepath


    sht = workbook.sheets['Sheet1'] #默认工作表 sheet1

    nrows = sht.used_range.last_cell.row #excel已使用总行数
    ncols = sht.used_range.last_cell.column
    '''expand(table)  读取a5开始所有表格中的二维数据，空值不读
        expand(right) 只读a5这一行
        #Data = sht.range((5, 1), (nrows - 9, 1)).expand('right').value
    '''

    title = sht.range('A1', 'H4').value  #读取表头
    data = sht.range(5, 1).expand('table').value
    datalen = len(data)
    endtitle = sht.range((nrows-8, 1), (nrows, 8)).value
    for i in range(0, datalen-1):# 总数据len条  循环次数1 +（len-1）次

        prj_name = data[i][1]
        try:
            result = class_text(prj_name)
        except:
            print('识别出错，请检查sheet名称和网络')
            workbook.close()
            sys.exit(1) #退出程序，1为异常退出，报错。 0为正常退出

        if result == '1':
            draw.append(data[i]) #识别结果为图结
        elif result == '0':
            ticket.append(data[i]) #识别结果为票结

    drawlen = len(draw)
    ticketlen = len(ticket)
    print('图结:')
    for i in range(drawlen):

        print(draw[i][1])
    print('票结:')
    for i in range(ticketlen):

        print(ticket[i][1])

#====================新建图结表格=========================

    DrawWb = xw.Book()
    DrawSht = DrawWb.sheets['sheet1'] #储存在新建工作表的sheet1

    DrawSht.range(1, 1).expand('table').value = title
    DrawSht.range(5, 1).expand('table').value = draw
    DrawSht.range(drawlen+5, 1).expand('table').value = endtitle  #表尾从正文数据的下一行开始写

    #fileName = os.path.split(filepath) #分割路径【0】和文件【1】
    path = LocalPath + '\\明信水印长滩图结' + filepath[-15:]  #filepath[-15:]获取字符串filepath右侧15个字符即日期
    DrawWb.save(path=path) #保存新建文件

#================新建票结表格============================
    TicketWb = xw.Book()
    TicketSht = TicketWb.sheets['sheet1']

    TicketSht.range(1, 1).expand('table').value = title
    TicketSht.range(5, 1).expand('table').value = ticket
    TicketSht.range(drawlen + 5, 1).expand('table').value = endtitle
    path = LocalPath + '\\明信水印长滩票结' + filepath[-15:]
    TicketWb.save(path=path)

    '''
    新建文件
    处理字符串命名
    读A1：H3
    读draw最后一行至末尾
    写表头表尾
    填写表尾金额和方量
    保存
    '''

    '''
    获取图片储存路径
    orc批量识别
    已识别移除
    根据分类写内容
    
    '''
    #关闭工作簿
    workbook.close()
    DrawWb.close()
    TicketWb.close()





'''
    wb.sheets['sheet1'].range('A2').options(transpose=True).value = [1,2,3] #写入列
    wb.sheets['sheet1'].range('A2').value = [1,2,3] #写入行
    wb.sheets['sheet1'].range('A2').options(expand='table').value= [[1,2],[3,4]] # 写入行
'''
