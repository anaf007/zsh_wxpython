"""
print 20%6
print "-------4----------4-------4"

print round(1.0/6)
print round(2.0/6)
print round(14.0/4)
print round(10.0/3)
print round(7.0/2)
print round(3.0/1)
print "-----------3----------3-------"
"""
print round(20/6.0)
print round(19/6.0)
print round(18/6.0)
print round(17/5.0)
print round(16/5.0)
print round(15/5.0)
print round(14/4.0)
print round(13/4.0)
print round(12/4.0)
print round(11/3.0)
print round(10/3.0)
print round(9/3.0)
print round(8/2.0)
print round(7/2.0)
print round(6/2.0)
print round(5/2.0)
print round(4/1.0)
print round(3/1.0)
print round(2/1.0)
print round(1/1.0)
index =20

for i in range(6,0,-1):
    if index-index/i==index:continue
    index =  index-index/i
    print str(index)+"--"




