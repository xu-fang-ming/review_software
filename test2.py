from operator import itemgetter
list1 = [{"num":3}, {"num":1}, {"num":5}, {"num":4}, {"num":2},{"num":6}]

row_by_num =sorted(list1,key=itemgetter("num"))

# print(row_by_num)

print(list1[0])

for i in row_by_num[0:5]:
    print(i)



