import glob
import os
from time import sleep
try:
    import docx
    import PyPDF2
except ModuleNotFoundError:
    print("Required dependencies not found!!")
    ip = input("Do you wish to install them?[y/n]")
    if(ip == "y" or ip == "Y"):
        os.system("pip install python-docx")
        os.system("pip install PyPDF2")


def getResultantFilename(s="New File", ext=""):
    temp = s[:]
    num = 1
    while(os.path.isfile(os.getcwd()+"/../OUTPUT/"+temp+".pdf")):
        temp = s[:]
        temp += "[{}]".format(num)
        num += 1
    if(ext):
        return temp+"."+ext
    return temp


def imageToPDFs(imageArr, filename):
    document = docx.Document()
    for i in imageArr:
        document.add_picture(i, width=docx.shared.Inches(6.5))
        document.add_page_break()
    document.save(filename)


def getImgArray():
    temp = []
    ImgArray = []
    temp.append(glob.glob('../DROP/*.jpg'))
    temp.append(glob.glob('../DROP/*.png'))
    temp.append(glob.glob('../DROP/*.jprg'))
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
    print(ip)
    if(ip[0] == "#"):
        newImgArray = ImgArray[:]
    else:
        for i in ip:
            try:
                newImgArray.append(ImgDict[i])
            except Exception:
                print("Invalid values!!")
                exit()
    print(newImgArray)
    # filename = getResultantFilename("test", "docx")
    # imageToPDFs(newImgArray, filename)


if __name__ == "__main__":
    # print(getImgArray())
    # main()
    imageArray = ['../DROP\\Screenshot (62).png', '../DROP\\Screenshot (63).png',
                  '../DROP\\Screenshot (64).png', '../DROP\\Screenshot (65).png', '../DROP\\Screenshot (66).png']
    imageToPDFs(imageArray, "test.docx")
