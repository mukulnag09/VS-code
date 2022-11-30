
from numbers import Integral


def su(s,k):
    p=set(s)
    index=[]
    for d in p:
        max=0
        for x in range(len(k)):
            if x in index:
                continue
            if s[x]==d:
                if max<k[x]:
                    max=k[x]
        print(k,d,p,s,max,k.index(max))
        index.append(k.index(max))
        k[k.index(max)]=0          
    return sum(k)

s=["n","y","w","e","k","w"]
k=[3,0,4,4,0,5]

print(su(s,k))