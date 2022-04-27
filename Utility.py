import os


def projectDirectory():

    projectDirectory=""

    fileDirectory=os.path.abspath(__file__)

    fileDirectoryArray=fileDirectory.split("\\")

    for i in fileDirectoryArray:

        projectDirectory=projectDirectory+"\\"+i
        if i=="File":
             break

    output=projectDirectory[1:]+"\\Output\\"

    print(output)

    return output
