#coding=utf-8
from xpinyin import Pinyin
__author__ = 'anngle'
data = [
[121.1, 0.02, 0.02,1.01, 4.04],
[122.1, 0.56, 0.6,1.01, 4.04],
[122.1, 3.79, 4.04,1.01, 4.04],
[123.1, 93.75, 100.0,1.01, 4.04],
[123.1, 0.01, 0.01,1.01, 4.04],
[124.1, 0.01, 0.01,1.01, 4.04],
[124.1, 1.01, 1.08,1.01, 4.04],
[124.1, 0.11, 0.11,1.01, 4.04],
[124.1, 0.05, 0.06,1.01, 4.04],
[125.1, 0.39, 0.41,1.01, 4.04],
]
from itertools import groupby
print [reduce(lambda x,y: [k, x[1]+y[1], x[2]+y[2],x[3]+y[3],x[4]+y[4]], rows) \
       for k, rows in groupby(data, lambda x: x[0])]

print Pinyin().get_pinyin(u"酒水王老吉250ml")