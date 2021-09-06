import os
os.system('cls')

def write_in_file(info, f_name):  # Writes list of lists into CSV file
    with open(f_name, "w") as f:
        for var in info:
            f.write("\n".join((",".join(var))))


def roll_no_func():                    # function which outputs the files corresponding to individual roll no.s
    subject_dict = {}
    DIRECTORY = "output_individual_roll"

    with open("regtable_old.csv", "r") as f:      #reading from regtable_old.csv file
        for var in f:
            var = var.strip().split(",")
            rollno, register_sem, _, subno, _, _, _, _, sub_type = var   
            if subno not in subject_dict:
                subject_dict[subno] = []
            subject_dict[subno].append([rollno, register_sem, subno, sub_type])     

    if not os.path.exists(DIRECTORY):    # checking for if a directory already exists
        os.makedirs(DIRECTORY)

    for subno in subject_dict:          
        write_in_file(subject_dict[subno], os.path.join(DIRECTORY, subno + ".csv"))     # calling write_in_file function


def subject_func():   # function which outputs the files containing data of each subject
    roll_dict = {}
    DIRECTORY = "output_by_subject"

    with open("regtable_old.csv", "r") as f:    #reading from regtable_old.csv file
        for var in f:
            var = var.strip().split(",")
            rollno, register_sem, _, subno, _, _, _, _, sub_type = var
            if rollno not in roll_dict:
                roll_dict[rollno] = []
            roll_dict[rollno].append([rollno, register_sem, subno, sub_type])

    if not os.path.exists(DIRECTORY):      # checking for if a directory already exists
        os.makedirs(DIRECTORY)

    for rollno in roll_dict:
        write_in_file(roll_dict[rollno], os.path.join(DIRECTORY, rollno + ".csv"))

roll_no_func()
subject_func()