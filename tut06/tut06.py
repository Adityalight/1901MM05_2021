import os
import re
import shutil

os.system("cls")
try:                               # using try and except block to create corrected_srt directory
    os.mkdir("./corrected_srt/")
except:
    print("corrected_srt directory already exsists.")


def regex_renamer():  # function to rename the web_series' files

    # Taking input from the user

    print("1. Breaking Bad")
    print("2. Game of Thrones")
    print("3. Lucifer")

    webseries_num = int(                              
        input("Enter the number of the web series that you wish to rename. 1/2/3: "))         #taking interger input from the user for webseries
    season_padding = int(input("Enter the Season Number Padding: "))                          #taking positive integer input from the user for season padding
    episode_padding = int(input("Enter the Episode Number Padding: "))                        #taking positive integer input from the user for episode padding

    if webseries_num == 1:                       #checking if the input for webseries number is 1
        src_dir = r"C:\Users\hp\Dropbox\PC\Documents\GitHub\1901MM05_2021\tut06\wrong_srt\Breaking Bad"          #source directory for copying files
        dest_dir = r"C:\Users\hp\Dropbox\PC\Documents\GitHub\1901MM05_2021\tut06\corrected_srt\Breaking Bad"     #destination directory
        shutil.copytree(src_dir, dest_dir)               #copying all the files from source to destination using shutil library
        files = os.listdir(dest_dir)                     #storing filenames in the form of a list
        spl1 = re.split('s01', files[0])                 #using regex to split the filename according to requirement
        a = '1'
        pad1 = a.zfill(season_padding)                   #storing season number in a variable according to season_padding input
        e = '1'
        epad1 = e.zfill(episode_padding)                 #storing episode number in a variable according to episode_padding input
        e = '2'
        epad2 = e.zfill(episode_padding)
        e = '3'
        epad3 = e.zfill(episode_padding)
        e = '4'
        epad4 = e.zfill(episode_padding)
        e = '5'
        epad5 = e.zfill(episode_padding)
        e = '6'
        epad6 = e.zfill(episode_padding)
        e = '7'
        epad7 = e.zfill(episode_padding)
        rename1 = spl1[0]+'Season '+str(pad1)+' Episode '+str(epad1)+'.mp4'        #storing the correct strings for renaming filenames
        rename2 = spl1[0]+'Season '+str(pad1)+' Episode '+str(epad2)+'.mp4'
        rename3 = spl1[0]+'Season '+str(pad1)+' Episode '+str(epad3)+'.mp4'
        rename4 = spl1[0]+'Season '+str(pad1)+' Episode '+str(epad4)+'.mp4'
        rename5 = spl1[0]+'Season '+str(pad1)+' Episode '+str(epad5)+'.mp4'
        rename6 = spl1[0]+'Season '+str(pad1)+' Episode '+str(epad6)+'.mp4'
        rename7 = spl1[0]+'Season '+str(pad1)+' Episode '+str(epad7)+'.mp4'
        rename8 = spl1[0]+'Season '+str(pad1)+' Episode '+str(epad1)+'.srt'
        rename9 = spl1[0]+'Season '+str(pad1)+' Episode '+str(epad2)+'.srt'
        rename10 = spl1[0]+'Season '+str(pad1)+' Episode '+str(epad3)+'.srt'
        rename11 = spl1[0]+'Season '+str(pad1)+' Episode '+str(epad4)+'.srt'
        rename12 = spl1[0]+'Season '+str(pad1)+' Episode '+str(epad5)+'.srt'
        rename13 = spl1[0]+'Season '+str(pad1)+' Episode '+str(epad6)+'.srt'
        rename14 = spl1[0]+'Season '+str(pad1)+' Episode '+str(epad7)+'.srt'
        os.chdir(   
            r"C:\Users\hp\Dropbox\PC\Documents\GitHub\1901MM05_2021\tut06\corrected_srt\Breaking Bad")    #changing current working directory
        os.rename(files[0], rename1)       #renaming filenames using os library
        os.rename(files[1], rename8)
        os.rename(files[2], rename2)
        os.rename(files[3], rename9)
        os.rename(files[4], rename3)
        os.rename(files[5], rename10)
        os.rename(files[6], rename4)
        os.rename(files[7], rename11)
        os.rename(files[8], rename5)
        os.rename(files[9], rename12)
        os.rename(files[10], rename6)
        os.rename(files[11], rename13)
        os.rename(files[12], rename7)
        os.rename(files[13], rename14)

    elif webseries_num == 2:                  #checking if the input for webseries number is 2
        src_dir = r"C:\Users\hp\Dropbox\PC\Documents\GitHub\1901MM05_2021\tut06\wrong_srt\Game of Thrones"            #source directory for copying files
        dest_dir = r"C:\Users\hp\Dropbox\PC\Documents\GitHub\1901MM05_2021\tut06\corrected_srt\Game of Thrones"       #destination directory
        shutil.copytree(src_dir, dest_dir)                        #copying all the files from source to destination using shutil library
        files = os.listdir(dest_dir)                              #storing filenames in the form of a list
        a = '8'
        pad1 = a.zfill(season_padding)                            #storing season number in a variable according to season_padding input
        e = '1'
        epad1 = e.zfill(episode_padding)                          #storing episode number in a variable according to episode_padding input
        e = '2'
        epad2 = e.zfill(episode_padding)
        e = '3'
        epad3 = e.zfill(episode_padding)
        e = '4'
        epad4 = e.zfill(episode_padding)
        e = '5'
        epad5 = e.zfill(episode_padding)
        e = '6'
        epad6 = e.zfill(episode_padding)
        spl1 = re.split('.WEB.REPACK.MEMENTO.en.mp4', files[0])   #using regex to split the filenames according to requirement
        spl2 = re.split('8x01', spl1[0])
        spl3 = re.split('.WEB.REPACK.MEMENTO.en.mp4', files[2])
        spl4 = re.split('8x02', spl3[0])
        spl5 = re.split('.WEB.REPACK.MEMENTO.en.mp4', files[4])
        spl6 = re.split('8x03', spl5[0])
        spl7 = re.split('.WEB.REPACK.MEMENTO.en.mp4', files[6])
        spl8 = re.split('8x04', spl7[0])
        spl9 = re.split('.WEB.REPACK.MEMENTO.en.mp4', files[8])
        spl10 = re.split('8x05', spl9[0])
        spl11 = re.split('.WEB.REPACK.MEMENTO.en.mp4', files[10])
        spl12 = re.split('8x06', spl11[0])
        rename1 = spl2[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad1)+spl2[1]+'.mp4'              #storing the correct strings for renaming filenames
        rename2 = spl2[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad1)+spl2[1]+'.str'
        rename3 = spl4[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad2)+spl4[1]+'.mp4'
        rename4 = spl4[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad2)+spl4[1]+'.str'
        rename5 = spl6[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad3)+spl6[1]+'.mp4'
        rename6 = spl6[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad3)+spl6[1]+'.str'
        rename7 = spl8[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad4)+spl8[1]+'.mp4'
        rename8 = spl8[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad4)+spl8[1]+'.str'
        rename9 = spl10[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad5)+spl10[1]+'.mp4'
        rename10 = spl10[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad5)+spl10[1]+'.str'
        rename11 = spl12[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad6)+spl12[1]+'.mp4'
        rename12 = spl12[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad6)+spl12[1]+'.str'
        os.chdir(
            r"C:\Users\hp\Dropbox\PC\Documents\GitHub\1901MM05_2021\tut06\corrected_srt\Game of Thrones")    #changing current working directory
        os.rename(files[0], rename1)            #renaming filenames using os library
        os.rename(files[1], rename2)
        os.rename(files[2], rename3)
        os.rename(files[3], rename4)
        os.rename(files[4], rename5)
        os.rename(files[5], rename6)
        os.rename(files[6], rename7)
        os.rename(files[7], rename8)
        os.rename(files[8], rename9)
        os.rename(files[9], rename10)
        os.rename(files[10], rename11)
        os.rename(files[11], rename12)
    elif webseries_num == 3:                 #checking if the input for webseries number is 3
        src_dir = r"C:\Users\hp\Dropbox\PC\Documents\GitHub\1901MM05_2021\tut06\wrong_srt\Lucifer"        #source directory for copying files
        dest_dir = r"C:\Users\hp\Dropbox\PC\Documents\GitHub\1901MM05_2021\tut06\corrected_srt\Lucifer"   #destination directory
        shutil.copytree(src_dir, dest_dir)                #copying all the files from source to destination using shutil library
        files = os.listdir(dest_dir)                      #storing filenames in the form of a list
        a = '6'
        pad1 = a.zfill(season_padding)                    #storing season number in a variable according to season_padding input
        e = '1'
        epad1 = e.zfill(episode_padding)                  #storing episode number in a variable according to episode_padding input
        e = '2'
        epad2 = e.zfill(episode_padding)
        e = '3'
        epad3 = e.zfill(episode_padding)
        e = '4'
        epad4 = e.zfill(episode_padding)
        e = '5'
        epad5 = e.zfill(episode_padding)
        e = '6'
        epad6 = e.zfill(episode_padding)
        e = '7'
        epad7 = e.zfill(episode_padding)
        e = '8'
        epad8 = e.zfill(episode_padding)
        e = '9'
        epad9 = e.zfill(episode_padding)
        e = '10'
        epad10 = e.zfill(episode_padding)
        spl1 = re.split('.HDTV.CAKES.en.mp4', files[0])        #using regex to split the filenames according to requirement
        spl2 = re.split('6x01', spl1[0])
        spl3 = re.split('.HDTV.CAKES.en.mp4', files[2])
        spl4 = re.split('6x02', spl3[0])
        spl5 = re.split('.HDTV.CAKES.en.mp4', files[4])
        spl6 = re.split('6x03', spl5[0])
        spl7 = re.split('.HDTV.CAKES.en.mp4', files[6])
        spl8 = re.split('6x04', spl7[0])
        spl9 = re.split('.HDTV.CAKES.en.mp4', files[8])
        spl10 = re.split('6x05', spl9[0])
        spl11 = re.split('.HDTV.CAKES.en.mp4', files[10])
        spl12 = re.split('6x06', spl11[0])
        spl13 = re.split('.HDTV.CAKES.en.mp4', files[12])
        spl14 = re.split('6x07', spl13[0])
        spl15 = re.split('.HDTV.CAKES.en.mp4', files[14])
        spl16 = re.split('6x08', spl15[0])
        spl17 = re.split('.HDTV.CAKES.en.mp4', files[16])
        spl18 = re.split('6x09', spl17[0])
        spl19 = re.split('.HDTV.CAKES.en.mp4', files[18])
        spl20 = re.split('6x10', spl19[0])
        rename1 = spl2[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad1)+spl2[1]+'.mp4'       #storing the correct strings for renaming filenames
        rename2 = spl2[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad1)+spl2[1]+'.str'
        rename3 = spl4[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad2)+spl4[1]+'.mp4'
        rename4 = spl4[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad2)+spl4[1]+'.str'
        rename5 = spl6[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad3)+spl6[1]+'.mp4'
        rename6 = spl6[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad3)+spl6[1]+'.str'
        rename7 = spl8[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad4)+spl8[1]+'.mp4'
        rename8 = spl8[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad4)+spl8[1]+'.str'
        rename9 = spl10[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad5)+spl10[1]+'.mp4'
        rename10 = spl10[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad5)+spl10[1]+'.str'
        rename11 = spl12[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad6)+spl12[1]+'.mp4'
        rename12 = spl12[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad6)+spl12[1]+'.str'
        rename13 = spl14[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad7)+spl14[1]+'.mp4'
        rename14 = spl14[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad7)+spl14[1]+'.str'
        rename15 = spl16[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad8)+spl16[1]+'.mp4'
        rename16 = spl16[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad8)+spl16[1]+'.str'
        rename17 = spl18[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad9)+spl18[1]+'.mp4'
        rename18 = spl18[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad9)+spl18[1]+'.str'
        rename19 = spl20[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad10)+spl20[1]+'.mp4'
        rename20 = spl20[0]+'Season ' + \
            str(pad1)+' Episode '+str(epad10)+spl20[1]+'.str'
        os.chdir(
            r"C:\Users\hp\Dropbox\PC\Documents\GitHub\1901MM05_2021\tut06\corrected_srt\Lucifer")    #changing current working directory
        os.rename(files[0], rename1)               #renaming filenames using os library
        os.rename(files[1], rename2)
        os.rename(files[2], rename3)
        os.rename(files[3], rename4)
        os.rename(files[4], rename5)
        os.rename(files[5], rename6)
        os.rename(files[6], rename7)
        os.rename(files[7], rename8)
        os.rename(files[8], rename9)
        os.rename(files[9], rename10)
        os.rename(files[10], rename11)
        os.rename(files[11], rename12)
        os.rename(files[12], rename13)
        os.rename(files[13], rename14)
        os.rename(files[14], rename15)
        os.rename(files[15], rename16)
        os.rename(files[16], rename17)
        os.rename(files[17], rename18)
        os.rename(files[18], rename19)
        os.rename(files[19], rename20)
    else:
        print("Wrong input given by user for webseries number.")

#calling the function to rename webseries filenames
regex_renamer()
