import os
os.system("cls")


def meraki_helper(n):
    meraki_num_sum = 0
    nonmeraki_num_sum=0
    for num in n:
        if((num//10) == 0):
            print("Yes -", num, " is a Meraki number")
            meraki_num_sum += 1
        else:
            flag=1
            n1=num
            dig1=n1%10
            n1=n1//10
            while(n1>0):
                dig2=n1%10
                if(abs(dig1-dig2)!=1):
                    flag=0
                    nonmeraki_num_sum+=1
                    break
                else:
                    n1=n1//10
                    dig1=dig2
            if(flag==1):
                print("Yes -",num," is a Meraki number")
                meraki_num_sum+=1
            if(flag==0):
                print("No -",num," is not a Meraki number")
    print("The input list contains ",meraki_num_sum,"meraki and ",nonmeraki_num_sum,"non meraki numbers")
    return


meraki_helper(n=[12, 14, 56, 78, 98, 54, 678, 134, 789, 0, 7, 5, 123, 45, 76345, 987654321])
