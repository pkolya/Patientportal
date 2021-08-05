str="Welcome to mindtree to welcome to mindtree"
str1=str.lower()
words=str1.split(" ")
countWelcome=0;
countMindtree=0;
for i in range(0,len(words)):
    if words[i]=="welcome":
        countWelcome +=1
    elif  words[i] =="mindtree":
        countMindtree +=1


print("Welcome=",countWelcome)
print("mindtree=",countMindtree)