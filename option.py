"""BY-冷月长空 QQ 1666888816"""
import xlrd   #pip install xlrd
import xlwt   #pip install xlrt
import re

####################配置区域#################

xls_path = '近代史选择题.xls'#当前目录下的文件名称
da=6                       #答案对应的列数---就是答案是ABCD的那一行|答案不是文字
qu=1                       #问题对应的列数

A=2                        #A选项对应的列数
B=3                        #B选项对应的列数
C=4                        #C选项对应的列数
D=5                        #D选项对应的列数
#E=9                       #E选项对应的列数


"""
你只需要将你的xls文件放在py目录下然后配置上方文件，运行py会自动生成一份新的xls
新的xls中问题和答案都是以文字来呈现

"""
####################配置区域#################


def main(xls_path):
    work_book = xlrd.open_workbook(xls_path)
    workbook = xlwt.Workbook(encoding = 'utf8')
    worksheet = workbook.add_sheet('sheet')
    worksheet.write(0, 0, "quesion")
    worksheet.write(0, 1, "answer")

    if not work_book:
        print('路径错误')
        return 0

    sheet = work_book.sheet_by_index(0) # 根据索引来获取sheet对象
    sheet_load = work_book.sheet_loaded(sheet_name_or_index=0)

    if not sheet_load:
        print('xlsx内容出错')
        return 0

    rows = sheet.nrows # 获取有效行数
    print("总共有",rows,"行")

    for row in range(rows):
        row += 1
        answer_list = []
        question = sheet.cell_value(row,qu-1)
        idxs = sheet.cell_value(row,da-1).replace('答案：','')
        if idxs:
            for i in idxs:
                if i == 'A':
                    answer_list.append(sheet.cell_value(row, A-1))
                if i == 'B':
                    answer_list.append(sheet.cell_value(row, B-1))

                if i == 'C':
                    answer_list.append(sheet.cell_value(row, C-1))

                if i == 'D':
                    answer_list.append(sheet.cell_value(row, D-1))
                    
                if i == 'E':
                    answer_list.append(sheet.cell_value(row, E-1))

            if len(answer_list)>1:
                answer = '#'.join(answer_list)
            else:
                answer = answer_list[0]

        print("正在处理第",row+1,"行","问题是",question,"答案是",answer)
        worksheet.write(row, 0, question)
        worksheet.write(row, 1, answer)
        workbook.save(xls_path[0:xls_path.rfind('.')] + '_new.xls')  # 保存工作簿



            
if __name__ == "__main__":
    main(xls_path)

