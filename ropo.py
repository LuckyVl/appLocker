import shutil, re, os

list1 = []
list2 = []
# считываем адреса откуда и куда
with open("resurse.txt") as file:
    for item in file:
        list1.append(item)
# print(list1)
file.close()
# чистим адреса
for i in list1:
    i = re.sub(r"\n","",i)
    list2.append(i)
# print(list2)
pathout = list2[0]
# print(pathout)
temp = pathout.split("\\")
# print(temp)
pathin = list2[1] + temp[-2]
# print(pathin)
shutil.copytree(pathout,pathin)