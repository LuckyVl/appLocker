import shutil, re

list1 = []
list2 = []
with open("resurse.txt") as file:
    for item in file:
        list1.append(item)
# print(list1)
file.close()
for i in list1:
    i = re.sub(r"\n","",i)
    list2.append(i)
# print(list2)
shutil.copy(str(list2[0]),str(list2[1]))