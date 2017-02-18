def grabElementList(lOfL, i):
    l=[]
    for element in lOfL:
        print("3:"+str(element))
        l.append(element[i])
    return l

def sumList(l):
    sum=0
    for element in l:
        print("2:"+str(element))
        sum+=element
    return sum

def multiply(m1, m2):
    m3=[[0]*3]*4
    print("m3: "+str(m3))
    for i in range(len(m1)-1):
        for j in range(len(m2[i])-1):
            m3[i][j]=(sumList(m1[i])*sumList(grabElementList(m2, j)))
    return m3

m1=[[1, 2, 3], [4, 5, 6], [7, 8, 9]]
m2=[[10, 11, 12], [13, 14, 15], [16, 17, 18]]

print(multiply(m1, m2))