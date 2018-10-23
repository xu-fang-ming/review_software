# def dec_per(rate):
#     b = '%.2f%%' % (rate * 100)
#     return b
#
# a=0.12898080
#
# c= dec_per(a)
#
# print(c)
#
# print(type(c))

# import re
# # # b = "1489859.8"
# # # lis = re.findall(r'.{3}', b)
# # # c = ','.join(lis)
# # #
# # # # print(b[0:-3])
# # #
# # # print(c)

# d = "-,123,111,222.89"
#
# if d.startswith("-") and d[1] == ",":
#     d = d.replace(",", "", 1)
#
# print(d)


rate = 0.12
print('%.f%%' % (rate * 100))