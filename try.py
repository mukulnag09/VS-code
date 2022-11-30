# Write any import statements here
def getMinimumDeflatedDiscCount(arr) -> int:

  # Write your code here
  dic={}
  for x in arr:
    if x in dic:
      dic[x]+=1
    else:
      dic[x]=1
  sort=sorted(dic.items(),key=lambda x:x[1])
  sort.reverse()
  count=0
  till=0
  for x in sort:
    till=x[1]+till
    count+=1
    if till>=len(arr)/2:
      return count





arr = [7,7,7,7,7,7]


print(getMinimumDeflatedDiscCount(arr))