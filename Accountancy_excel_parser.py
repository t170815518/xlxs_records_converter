'''
读取公司记账凭证xls总文件，并为每对借贷方生成单独xls文件供打印使用
'''

# 读取需要的库：xlwt用于写xls文件，csv用于读取表头并提取时间
import pandas as pd
import xlwt
import numpy as np


def LenderChecker(row):
    '''根据row判断此记录是否为借方，是则返回True'''
    if row['借方'] is not np.nan and row['贷方'] is np.nan:
        return True
    else:
        return False


def DateIdenticalChecker(series):
    '''比较series里的值是否相同，是则返回True'''
    a = series.to_numpy()
    return (a[0] == a).all()


class ExcelSheet:
    '''即将导出的Excel表格对象
        attributes:
        time: the sheet's time in string
        lenders: the list of Lender
        borrowers: the list of Borrower'''
    # 分界线格式
    border = xlwt.Borders()
    border.left = xlwt.Borders.THIN
    border.right = xlwt.Borders.THIN
    border.top = xlwt.Borders.THIN
    border.bottom = xlwt.Borders.THIN
    # 凭证标题字体
    titleFont = xlwt.Font()
    titleFont.bold = True
    titleFont.height = 18*20
    # 列标题字体
    colFont = xlwt.Font()
    colFont.bold = True
    titleFont.height = 12*20
    # 居中
    center_alignment = xlwt.Alignment()
    center_alignment.horz = center_alignment.HORZ_CENTER
    center_alignment.vert = center_alignment.VERT_CENTER

    def __init__(self, index, time, lenders, borrowers, generated_workbook):
        self.id = index
        self.time = time
        self.lenders = lenders
        self.borrows = borrowers
        self.sum_lend = 0
        self.sum_borrow = 0
        self.workbook = generated_workbook
        for x in lenders:
            self.sum_lend += x.money
        for x in borrowers:
            self.sum_borrow += x.money

    def to_xls(self):
        '''生成xls文件'''
        ws = self.workbook.add_sheet('{}'.format(round(self.id)))
        row = self.heading_format(ws)
        row = self.body_format(ws, row)
        self.ending_format(ws, row)

    def heading_format(self, ws):
        '''调整xls表头格式
        :argument
        ws: xlrd.sheet object
        :return row index to begin body writing
        '''
        # 标题格式
        headerFormat = xlwt.XFStyle()
        headerFormat.font = ExcelSheet.titleFont
        headerFormat.alignment = ExcelSheet.center_alignment
        # 副标题格式
        subHeaderFormat = xlwt.XFStyle()
        subHeaderFormat.alignment = ExcelSheet.center_alignment

        ws.write_merge(0, 0, 0, 4, "记账凭证", headerFormat)
        ws.write_merge(1, 1, 0, 4, "日期: "+self.time+"\t凭证记账号：{}          附件：      张".format(round(self.id)), subHeaderFormat)
        # 表列标题格式
        colHeaderFormat = xlwt.XFStyle()
        colHeaderFormat.font = ExcelSheet.colFont
        colHeaderFormat.borders = ExcelSheet.border
        ws.write_merge(2, 3, 0, 0, "摘要", colHeaderFormat)
        ws.write_merge(2, 2, 1, 2, "会计科目", colHeaderFormat)
        ws.write_merge(2, 2, 3, 4, "金额", colHeaderFormat)
        ws.write(3, 1, "一级科目", colHeaderFormat)
        ws.write(3, 2, "二级科目", colHeaderFormat)
        ws.write(3, 3, "借", colHeaderFormat)
        ws.write(3, 4, "贷", colHeaderFormat)
        return 4

    def body_format(self,ws,StartRowId):
        bodyFormat = xlwt.XFStyle()
        bodyFormat.borders = ExcelSheet.border
        bodyFormat.alignment = ExcelSheet.center_alignment
        bodyFormat.alignment.wrap = 1

        row = StartRowId
        for x in self.lenders:
            ws.write(row, 0, x.abs, bodyFormat)
            cwidth = ws.col(0).width
            ws.write(row, 1, x.cate1, bodyFormat)
            if x.cate2 is not np.nan:
                ws.write(row, 2, x.cate2, bodyFormat)
            else:
                ws.write(row, 2, " ", bodyFormat)
            ws.write(row, 3, x.money, bodyFormat)
            ws.write(row, 4, " ", bodyFormat)  # write the empty cell to make the broader complete
            row += 1

        for x in self.borrows:
            ws.write(row, 0, x.abs, bodyFormat)
            ws.write(row, 1, x.cate1, bodyFormat)
            if x.cate2 is not np.nan:
                ws.write(row, 2, x.cate2, bodyFormat)
            else:
                ws.write(row, 2, " ", bodyFormat)
            ws.write(row, 4, x.money, bodyFormat)
            ws.write(row, 3, " ", bodyFormat)  # write the empty cell to make the broader complete
            row += 1

        # adjust column width
        ws.col(0).width = round(2 * ws.col(0).width)
        ws.col(1).width = round(1.15 * ws.col(1).width)
        ws.col(2).width = round(1.15 * ws.col(2).width)

        return row

    def ending_format(self,ws, StartRowId):
        colHeaderFormat = xlwt.XFStyle()
        colHeaderFormat.font = ExcelSheet.colFont
        colHeaderFormat.borders = ExcelSheet.border

        otherFormat = xlwt.XFStyle()
        otherFormat.borders = ExcelSheet.border

        row = StartRowId
        ws.write(row, 0, "总计：", colHeaderFormat)
        for i in range(1,3):
            ws.write(row, i, " ", otherFormat)  # write the empty cell to make the broader complete
        ws.write(row, 3, self.sum_lend, otherFormat)
        ws.write(row, 4, self.sum_borrow, otherFormat)

        row += 1
        ws.write_merge(row, row, 0, 4, "  主管：                复核：                记账：                制单：                ")


class Lender:
    '''借方对象'''
    def __init__(self,abstract,cate1,money,cate2=None):
        self.abs = abstract
        self.cate1 = cate1
        self.cate2 = cate2
        self.money = money


class Borrower:
    '''贷方对象'''
    def __init__(self,abstract,cate1,money,cate2=None):
        self.abs = abstract
        self.cate1 = cate1
        self.cate2 = cate2
        self.money = money


FileName = 'example.xls'  # 总文件路径
sheet_names = pd.ExcelFile(FileName).sheet_names

for sheetName in sheet_names:  # 遍历所有sheet
    try:
        overall = pd.read_excel(FileName, header=1, usecols=['凭证号', '月', '日', '凭证摘要', '借方', '二级明细', '金额', '贷方',
                                                             '二级明细.1', '金额.1'], sheet_name=sheetName)

        MaxIndex = int(overall['凭证号'].max())

        GroupByIndex = overall.groupby('凭证号')
    except KeyError:
        print("KeyError: {} sheet 无法被处理".format(sheetName))
        continue

    generated = xlwt.Workbook()

    for index, group in GroupByIndex:
        Borrowers = []
        Lenders = []
        for _, row in group.iterrows():
            if LenderChecker(row):
                Lenders.append(Lender(abstract=row['凭证摘要'],cate1=row['借方'],cate2=row['二级明细'],money=row['金额']))
            else:
                Borrowers.append(Borrower(abstract=row['凭证摘要'],cate1=row['贷方'],cate2=row['二级明细.1'],money=row['金额.1']))
        if DateIdenticalChecker(group['日']) and DateIdenticalChecker(group['月']):
            TimeString = str(int(row["月"]))+"月"+str(int(row["日"]))+"日"
        else:
            print("警告：检测到序号{}的记录日期不相同，默认选择序号中第一行记录日期，请稍后手动修改".format(index))

        sum_borrow = 0
        for x in Borrowers:
            sum_borrow += x.money
        sum_lend = 0
        for x in Lenders:
            sum_lend += x.money
        if sum_lend != sum_borrow:
            print("警告：检测到序号{}的金额不匹配——借方金额={},贷方金额={}，请稍后检查手动修改".format(index, sum_lend, sum_borrow))

        ExcelSheet(index, TimeString, Lenders, Borrowers, generated).to_xls()

    generated.save(sheetName + '.xls')
    print("{}的凭证已生成".format(sheetName))
print("========程序完成========")

