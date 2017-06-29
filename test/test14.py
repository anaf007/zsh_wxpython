import math
users =['a','b','c','d','e','f']
sku = ['1','2','3','4','5','6','7','8']
data_list = []
for u in range(len(users),0,-1):
    for l in range(len(sku),len(sku)-\
            int(math.ceil(len(sku)/float(u))),-1):
        data_list.append([sku[l-1],users[u-1]])
        sku.pop()
for x in data_list:
    print x

