import docx
from openpyxl import load_workbook
import re

# 获取文档对象
file_path_notes = r"D:\data\附注\报告-中瑞诚0910.docx"
file_notes = docx.Document(file_path_notes)

# print("表格数:" + str(len(file.tables)))
#
# num = 0
# for i in file_notes.tables:
#     print(str(num)+i.cell(1,0).text)
#     num += 1

print(file_notes.tables[7].cell(1,1).text)
print("---------------")
print(file_notes.tables[7].cell(1,2).text)
# 第二次找到中间表
file_path_mid = r"D:\data\中间表\输出中间表8.xlsx"
wb_mid = load_workbook(filename=file_path_mid,data_only=True)
sheets_mid = wb_mid.sheetnames
sheet_first_mid = sheets_mid[0]  # 中间表1
ws_mid = wb_mid[sheet_first_mid]  # 中间表工作区1

# 中间附注表二
sheet_second_mid = sheets_mid[1]
ws_mid_second = wb_mid[sheet_second_mid]


# ######################开始往附注中填数###########################

def standard_num(x):
    """
    :param x:需要标准化的数字
    :return: 标准化后的数字
    """
    if x == "None":
        return "0.00"
    else:
        if '.' in x:
            num = len(x) % 3
            if num == 0:
                n = x[0:-3]
            elif num == 1:
                n = "00" + x[0:-3]
            elif num == 2:
                n = "0" + x[0:-3]
            lis = re.findall(r'.{3}', n)
            c = ','.join(lis)

            c = c + x[-3:]
        else:
            num = len(x) % 3
            if num == 0:
                n = x
            elif num == 1:
                n = "00" + x
            elif num == 2:
                n = "0" + x

            lis = re.findall(r'.{3}', n)
            c = ','.join(lis)
            c = c+".00"
        return c.lstrip("0")


# 1.货币资金
hb1 = standard_num(str(ws_mid["C8"].value))
hb2 = standard_num(str(ws_mid["C9"].value))
hb3 = standard_num(str(ws_mid["C10"].value))
hb4 = standard_num(str(ws_mid["C11"].value))
hb5 = standard_num(str(ws_mid["C14"].value))
hb6 = standard_num(str(ws_mid["C15"].value))
hb7 = standard_num(str(ws_mid["C16"].value))
hb8 = standard_num(str(ws_mid["C17"].value))

file_notes.tables[2].cell(1, 1).text = hb1
file_notes.tables[2].cell(2, 1).text = hb2
file_notes.tables[2].cell(3, 1).text = hb3
file_notes.tables[2].cell(4, 1).text = hb4
file_notes.tables[2].cell(1, 2).text = hb5
file_notes.tables[2].cell(2, 2).text = hb6
file_notes.tables[2].cell(3, 2).text = hb7
file_notes.tables[2].cell(4, 2).text = hb8

# 2.以公允价值计量且其变动计入当期损益的金融资产
gy1 = standard_num(str(ws_mid["C23"].value))
gy2 = standard_num(str(ws_mid["C24"].value))
gy3 = standard_num(str(ws_mid["C25"].value))
gy4 = standard_num(str(ws_mid["C28"].value))
gy5 = standard_num(str(ws_mid["C32"].value))
gy6 = standard_num(str(ws_mid["C34"].value))
gy7 = standard_num(str(ws_mid["C35"].value))
gy8 = standard_num(str(ws_mid["C36"].value))
gy9 = standard_num(str(ws_mid["C39"].value))
gy10 = standard_num(str(ws_mid["C43"].value))

file_notes.tables[3].cell(1, 1).text = gy1
file_notes.tables[3].cell(2, 1).text = gy2
file_notes.tables[3].cell(3, 1).text = gy3
file_notes.tables[3].cell(4, 1).text = gy4
file_notes.tables[3].cell(6, 1).text = gy5
file_notes.tables[3].cell(1, 2).text = gy6
file_notes.tables[3].cell(2, 2).text = gy7
file_notes.tables[3].cell(3, 2).text = gy8
file_notes.tables[3].cell(4, 2).text = gy9
file_notes.tables[3].cell(6, 2).text = gy10

# 3.应收票据及应收账款
pj_zk1 = standard_num(str(ws_mid["C51"].value))
pj_zk2 = standard_num(str(ws_mid["C78"].value))
pj_zk3 = str(float(pj_zk1)+float(pj_zk2))
pj_zk4 = standard_num(str(ws_mid["C55"].value))
pj_zk5 = standard_num(str(ws_mid["C106"].value))
pj_zk6 = str(float(pj_zk4)+float(pj_zk5))

file_notes.tables[4].cell(1, 1).text = pj_zk1
file_notes.tables[4].cell(2, 1).text = pj_zk2
file_notes.tables[4].cell(3, 1).text = pj_zk3
file_notes.tables[4].cell(1, 2).text = pj_zk4
file_notes.tables[4].cell(2, 2).text = pj_zk5
file_notes.tables[4].cell(3, 2).text = pj_zk6


# 3.1应收票据
pj1 = standard_num(str(ws_mid["C49"].value))
pj2 = standard_num(str(ws_mid["C50"].value))
pj3 = standard_num(str(ws_mid["C51"].value))
pj4 = standard_num(str(ws_mid["C53"].value))
pj5 = standard_num(str(ws_mid["C54"].value))
pj6 = standard_num(str(ws_mid["C55"].value))

file_notes.tables[5].cell(1, 1).text = pj1
file_notes.tables[5].cell(2, 1).text = pj2
file_notes.tables[5].cell(3, 1).text = pj3
file_notes.tables[5].cell(1, 2).text = pj4
file_notes.tables[5].cell(2, 2).text = pj5
file_notes.tables[5].cell(3, 2).text = pj6

# 3.2应收账款(期末数)
zk1 = standard_num(str(ws_mid["C72"].value))
zk2 = standard_num(str(ws_mid["C73"].value))
zk3 = standard_num(str(ws_mid["C74"].value))
zk4 = standard_num(str(ws_mid["C75"].value))
zk5 = standard_num(str(ws_mid["C78"].value))
if float(zk5) == 0:
    zk6 = "0.00"
    zk7 = "0.00"
    zk8 = "0.00"
    zk9 = "0.00"
else:
    zk6 = "%.2f%%" % (float(zk1)/float(zk5))
    zk7 = "%.2f%%" % (float(zk2)/float(zk5))
    zk8 = "%.2f%%" % (float(zk3)/float(zk5))
    zk9 = "%.2f%%" % (float(zk4)/float(zk5))

file_notes.tables[6].cell(2, 1).text = zk1
file_notes.tables[6].cell(3, 1).text = zk2
file_notes.tables[6].cell(4, 1).text = zk3
file_notes.tables[6].cell(5, 1).text = zk4
file_notes.tables[6].cell(6, 1).text = zk5
file_notes.tables[6].cell(2, 2).text = zk6
file_notes.tables[6].cell(3, 2).text = zk7
file_notes.tables[6].cell(4, 2).text = zk8
file_notes.tables[6].cell(5, 2).text = zk9

# 3.2应收账款(期初数)
zk10 = standard_num(str(ws_mid["C100"].value))
zk11 = standard_num(str(ws_mid["C101"].value))
zk12 = standard_num(str(ws_mid["C102"].value))
zk13 = standard_num(str(ws_mid["C103"].value))
zk14 = standard_num(str(ws_mid["C106"].value))

if float(zk14) == 0:
    zk15 = "0.00"
    zk16 = "0.00"
    zk17 = "0.00"
    zk18 = "0.00"
else:
    zk15 = "%.2f%%" % (float(zk10)/float(zk14))
    zk16 = "%.2f%%" % (float(zk11)/float(zk14))
    zk17 = "%.2f%%" % (float(zk12)/float(zk14))
    zk18 = "%.2f%%" % (float(zk13)/float(zk14))

file_notes.tables[6].cell(2, 3).text = zk10
file_notes.tables[6].cell(3, 3).text = zk11
file_notes.tables[6].cell(4, 3).text = zk12
file_notes.tables[6].cell(5, 3).text = zk13
file_notes.tables[6].cell(6, 3).text = zk14
file_notes.tables[6].cell(2, 4).text = zk15
file_notes.tables[6].cell(3, 4).text = zk16
file_notes.tables[6].cell(4, 4).text = zk17
file_notes.tables[6].cell(5, 4).text = zk18

# 3.3 期末余额前五名
zk_qm1 = standard_num(str(ws_mid_second["A60"].value))
zk_qm2 = standard_num(str(ws_mid_second["A61"].value))
zk_qm3 = standard_num(str(ws_mid_second["A62"].value))
zk_qm4 = standard_num(str(ws_mid_second["A63"].value))
zk_qm5 = standard_num(str(ws_mid_second["A64"].value))
zk_qm6 = standard_num(str(ws_mid_second["B60"].value))
zk_qm7 = standard_num(str(ws_mid_second["B61"].value))
zk_qm8 = standard_num(str(ws_mid_second["B62"].value))
zk_qm9 = standard_num(str(ws_mid_second["B63"].value))
zk_qm10 = standard_num(str(ws_mid_second["B64"].value))
zk_qm11 = standard_num(str(ws_mid_second["B65"].value))
zk_qm12 = standard_num(str(ws_mid_second["C60"].value))
zk_qm13 = standard_num(str(ws_mid_second["C61"].value))
zk_qm14 = standard_num(str(ws_mid_second["C62"].value))
zk_qm15 = standard_num(str(ws_mid_second["C63"].value))
zk_qm16 = standard_num(str(ws_mid_second["C64"].value))
zk_qm17 = standard_num(str(ws_mid_second["C65"].value))

file_notes.tables[7].cell(3, 3).text = zk10
file_notes.tables[7].cell(4, 3).text = zk11
file_notes.tables[7].cell(5, 3).text = zk12
file_notes.tables[7].cell(6, 3).text = zk13
file_notes.tables[7].cell(7, 3).text = zk14
file_notes.tables[7].cell(3, 4).text = zk15
file_notes.tables[7].cell(4, 4).text = zk16
file_notes.tables[7].cell(5, 4).text = zk17
file_notes.tables[7].cell(6, 4).text = zk18
file_notes.tables[7].cell(3, 3).text = zk10
file_notes.tables[7].cell(4, 3).text = zk11
file_notes.tables[7].cell(5, 3).text = zk12
file_notes.tables[7].cell(6, 3).text = zk13
file_notes.tables[7].cell(7, 3).text = zk14
file_notes.tables[7].cell(3, 4).text = zk15
file_notes.tables[7].cell(4, 4).text = zk16
file_notes.tables[7].cell(5, 4).text = zk17
file_notes.tables[7].cell(6, 4).text = zk18


