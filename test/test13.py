sku =20
temp = sku
bk =0
cn=0
for u in range(6,0,-1):
    # if index-index/i==index:continue
    cn =bk
    sku =  sku-sku/u
    bk = temp-sku-1
    for s in range(sku):
        if bk<s:break
        if cn>=s:continue
        print  str(s)+"-"+str(u)
    print str(cn)+"---------"+str(bk)
