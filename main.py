import os
import time

try:
    from docx import *
except:
    print("first install docx library")

bluh = str(input("""Hi!
this program recives an string and find that that string in all docx files and cheange them to what you want

--before pressing Enter make sure you closed the folder that you want to change the files in that--"""))

old_text = str(input("Enter the string you want to find and delete: "))
new_text = str(input("Enter the string you want to replace with deleted string: "))
main_path = str(input("Enter the main folder path: "))
main_path = main_path.replace("\\", '/')

def get_all(path):
    global paths
    paths = []
    for root, dirs, files in os.walk(path):
        for file in files:
            paths.append(os.path.join(root, file))
    return paths

def change(name):
    doc = Document(r'%s' % name)

    for para in doc.paragraphs:
            if old_text in para.text:
                para.text = para.text.replace(old_text, new_text)

    doc.save(r'%s' % name)


get_all(main_path)
try:
    for p in paths:
        change(p)
except:
    print("we can not to that it may be beacuse of")
    print("maybe the folder is open!")
    print("make sure to close it before you try to run this file")

print("Done!")
print("bye")
time.sleep(2)

