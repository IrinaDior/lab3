import os
dataList = []
marketList = []
hit = False
groupStart = 0

try:
    fh = open('groups.txt', 'r')
    print("File opened.")
except FileNotFoundError:
    print("File is not found")

print("Groups has been imported");

data = fh.readline()

groupcounter = 0;
fh.seek(0)
for x in fh:
    #print(x)
    dataList.append(x)
    if "." in x:
        marketList.append(groupcounter)
        if "IKTE" in x and hit is False:
            groupStart = len(marketList)
            hit = True

    groupcounter = groupcounter + 1

print("Groups in the file: ")

for x in marketList:
    print(dataList[x])


print("Enter the filename for IKTE Export:")
filerec = input()
filerec = filerec + ".txt"

f = open(filerec, "w+")
for i in range(marketList[int(groupStart)-1], marketList[int(groupStart)]):
        print(dataList[i])
        f.write("%s\n" % dataList[i])
f.close()
print("Finish. Group is exported")
os.startfile(filerec)