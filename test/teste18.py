__author__ = 'anngle'
import datetime
starttime = datetime.datetime.now()
temp = 0
for i in range(100000000):
    temp = i



endtime = datetime.datetime.now()
print (endtime - starttime).seconds
