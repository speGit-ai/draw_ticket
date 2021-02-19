from Classifi import class_text
import xlwings as xw

draw = []
ticket = []

if __name__ == '__main__':

    a = "7#楼水池侧壁"
    #Classifi.class_text(a)

    excel = xw.App(visible=False, add_book=False)
    excel.display_alerts = False
    excel.screen_updating = False
    # 文件位置：filepath，打开test文档，然后保存，关闭，结束程序
    filepath = r'C:\Users\Administrator\Desktop\2021明信水印长滩12.25-1.25.xlsx'
    try:
        #workbook = xw.Book('2021明信水印长滩12.25-1.25.xlsx')

        workbook = xw.books.active
    except:
        workbook = excel.books.open(filepath)


    sht = workbook.sheets['sheet2']

    nrows = sht.used_range.last_cell.row
    ncols = sht.used_range.last_cell.column
    '''expand(table)  读取a5开始所有表格中的二维数据，空值不读
        expand(right) 只读a5这一行
        #Data = sht.range((5, 1), (nrows - 9, 1)).expand('right').value
    '''
    data = sht.range(5, 1).expand('table').value
    datalen = len(data)
    for i in range(0, datalen):

        prj_name = data[i][1]
        result = class_text(prj_name)
        if result == '1':
            draw.append(data[i])
        elif result == '0':
            ticket.append(data[i])

    drawlen = len(draw)
    ticketlen = len(ticket)
    print('图结:')
    for i in range(drawlen):

        print(draw[i][1])
    print('票结:')
    for i in range(ticketlen):

        print(ticket[i][1])

    DrawWb = xw.Book()
    DrawSht = DrawWb.sheets['sheet1']
    DrawSht.range(5, 1).expand('table').value = draw
    DrawWb.save(path=r'C:\Users\Administrator\Desktop\2021明信水印长滩图结12.25-1.25.xlsx')
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

    workbook.close()
    DrawWb.close()






'''
    wb.sheets['sheet1'].range('A2').options(transpose=True).value = [1,2,3] #写入列
    wb.sheets['sheet1'].range('A2').value = [1,2,3] #写入行
    wb.sheets['sheet1'].range('A2').options(expand='table').value= [[1,2],[3,4]] # 写入行
'''
