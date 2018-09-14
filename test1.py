# import re
# b = "11122233.02"
# b_len = len(b)
#
#
# c = ""
# if '.' in b:
#     num = b_len % 3
#     if num == 0:
#         n =b[0:-3]
#     elif num == 1:
#         n = "00"+b[0:-3]
#     elif num == 2:
#         n = "0" +b[0:-3]
#
#     lis = re.findall(r'.{3}', n)
#     print("b1:",lis)
#     c = ','.join(lis)
#     # print("n:",n)
#
#     c = c + b[-3:]
# else:
#     num = b_len % 3
#     if num == 0:
#         n =b
#     elif num == 1:
#         n = "00"+b
#     elif num == 2:
#         n = "0" +b
#
#     b = re.findall(r'.{3}', n)
#     print("b2:", b)
#     c = ','.join(b)
#     c = c+".00"
#
#
# print(c)
# print(c.lstrip("0"))


import re
b = "11122233.02"


def standard_num(x):
    """
    :param x:需要标准化的数字
    :return: 标准化后的数字
    """
    if '.' in x:
        num = len(x) % 3
        if num == 0:
            n =x[0:-3]
        elif num == 1:
            n = "00"+x[0:-3]
        elif num == 2:
            n = "0" +x[0:-3]
        lis = re.findall(r'.{3}', n)
        c = ','.join(lis)

        c = c + x[-3:]
    else:
        num = len(x) % 3
        if num == 0:
            n =x
        elif num == 1:
            n = "00"+x
        elif num == 2:
            n = "0" +x

        lis = re.findall(r'.{3}', n)
        c = ','.join(lis)
        c = c+".00"
    return c.lstrip("0")


a=standard_num(b)

print(a)


