import os

for cur_year in range (1,6):
    path = "C:\\test\\kurs"+str(cur_year)
    os.makedirs(path+"\\reference")

    print(path)
    open(path+"\\reference\\reference.html", 'a')
    open(path + "\\reference\\readme.txt", 'a')
    if cur_year is 1 or 3:
        os.makedirs(path + '\\install')
        os.chdir(path + '\\install')
        os.mkdir('C++')
        os.mkdir('web')
    for group in range(1,3):
        if cur_year is not 5:
            os.makedirs(path + "\\CS-" + str(cur_year) + ".1.0"+str(group))
        os.makedirs(path + "\\IK-" + str(cur_year) + ".1.0" + str(group))
        os.makedirs(path + "\\SE-" + str(cur_year) + ".1.0" + str(group))
