import shutil, re, os

list1 = []
list2 = []
print("Подготовка")
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
# print(temp[-1])
if temp[-1] == "":
    pathin = list2[1] + temp[-2]
else:
    pathin = list2[1] + temp[-1]
# print(pathin)
print("Копирование...")
# shutil.copytree(pathout,pathin)
print("Копирование выполнено")
