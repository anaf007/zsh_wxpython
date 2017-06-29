"""
a = [('L1-2', 'URLA'), (6, 'URLB'), (33, 'URLC'), ('L1-9', 'URLA'), (365, 'URLD'), ('B1-4', 'URLA'), (8, 'URLB')]
d = {}
for i, j in a:
    d[j] = str(d.get(j, 0)) + ','+str(i)
a = [(d[k], k) for k in d]
a.sort()
"""
# print a
acb = {'853':['G7-3',120],'854':['G7-5',125],'856':['G7-6',160]}
dd = {}
cc = [[853,'G7-3',120],[854,'G7-3',120],[854,'G1-3',120],\
[851,'G7-3',120],[853,'G8-3',20]]
for a,b,c in cc:
	if dd.get(str(a)):
		dd[str(a)] = [a,str(dd.get(str(a))[1]) + \
		','+str(b),dd.get(str(a))[2]+c]
	else:
		dd[str(a)] = [a,str(b),c]
aa = [(dd[k], k) for k in dd]
for k,v in aa:
	print k



 