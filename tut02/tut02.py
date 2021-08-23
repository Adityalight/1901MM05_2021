import os
os.system("cls")


def get_memory_score(input_nums):     # Function to calculate total score
    memory_list = []                    # list to store 5 previously called numbers
    score = 0
    for n in input_nums:
        if n in memory_list:          # checking for presence of each element of input list in memory list
            score += 1
        else:
            if len(memory_list) == 5:   # checking if memory list is full
                # removing the number stored for longest time in memory list
                memory_list.pop(0)
            memory_list.append(n)     # inserting new number in memory list
    print("Score: ", score)


input_nums = [3, 4, 1, 6, 3, 3, 9, 0, 0, 0]
invalid_inps = []                       # list to store invalid inputs
flag = 1
for num in input_nums:
    # checking if the elements in input list is a digit
    x = str(num).isdigit()
    if x == False:
        invalid_inps.append(num)      # adding invalid inputs to invalid list
        flag = 0
    else:
        continue
if flag == 1:
    # calling get_memory_score function when input list is valid
    get_memory_score(input_nums)
else:
    print("Please enter a valid input list. Invalid inputs detected: ", invalid_inps)
    exit(0)
