
num=int(input("enter a number"))
if num>0:
    for i in range(2,num):
        if(num%i==0):
            nums="not prime"
            break

        else:
            nums="prime"
    #else:
     #   print("invalid")

    print(nums)
6