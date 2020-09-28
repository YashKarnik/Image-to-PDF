import glob
import os
from datetime import datetime
from time import sleep
try:
    import docx
except ModuleNotFoundError:
    print("Required dependencies not found!!")
    ip = input("Do you wish to install them?[y/n]")
    if(ip == "y" or ip == "Y"):
        os.system("pip install python-docx")
    else:
        print("Goodbye")
        sleep(1)
        exit()


def getDeafultFilename():
    return "New Doc-"+datetime.now().strftime("%d-%m-%Y-%H:%M:%S")


def getResultantFilename(s=getDeafultFilename(), ext="docx"):
    temp = s[:]
    num = 1
    while(os.path.isfile(os.getcwd()+"/../OUTPUT/"+temp+"."+ext)):
        temp = s[:]
        temp += "[{}]".format(num)
        num += 1
    if(ext):
        return temp+"."+ext
    return temp.replace(" ", "-")


def setOrientation(document):
    curr_section = document.sections[-1]
    curr_section.left_margin = docx.shared.Inches(0.2)
    curr_section.right_margin = docx.shared.Inches(0.2)
    curr_section.top_margin = docx.shared.Inches(0.2)
    curr_section.bottom_margin = docx.shared.Inches(0.2)
    curr_section.orientation = docx.enum.section.WD_ORIENT.LANDSCAPE
    curr_section.page_width, curr_section.page_height = curr_section.page_height, curr_section.page_width
    img_width = curr_section.page_width-2*curr_section.left_margin
    img_height = curr_section.page_height-2*curr_section.bottom_margin
    return document, img_width, img_height


def imageToPDFs(imageArr, filename):
    document = docx.Document()
    document, img_width, img_height = setOrientation(document)
    for i in imageArr:
        document.add_picture(i, width=img_width, height=img_height)
    document.save("../OUTPUT/"+filename)


def getImgArray():
    temp = []
    ImgArray = []
    temp.append(glob.glob('../DROP/*.jpg'))
    temp.append(glob.glob('../DROP/*.png'))
    temp.append(glob.glob('../DROP/*.jpeg'))
    for i in temp:
        for j in i:
            ImgArray.append(j)
    return ImgArray


def main():
    ImgArray = getImgArray()
    newImgArray = []
    print(len(ImgArray), "Files found!!")
    ImgDict = {}
    if(len(ImgArray) == 0):
        print("No images found.Please drop images in DROP folder")
        exit()
    for i, j in enumerate(ImgArray):
        print(chr(i+97), ":", j.split("\\")[1])
        ImgDict[chr(i+97)] = j
    ip = input("Enter required order [#:default]: ").lower().strip().split()
    filename = input("Enter name of file(without extension):")
    filename = getResultantFilename(filename, "docx")
    if(ip[0] == "#"):
        newImgArray = ImgArray
    else:
        for i in ip:
            try:
                newImgArray.append(ImgDict[i])
            except Exception:
                print("Invalid values!!")
                exit()
    imageToPDFs(newImgArray, filename)
    print("DONE!")
    sleep(1)
    print("Goodbye!")
    sleep(1)


if __name__ == "__main__":
    try:
        os.mkdir(os.getcwd()+"/../OUTPUT")
    except FileExistsError as e:
        pass
    try:
        os.mkdir(os.getcwd()+"/../DROP")
    except FileExistsError as e:
        pass
    main()
