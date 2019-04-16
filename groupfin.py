import os
dataList = []
marketList = []
markerData = []
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
        markerData.append(x)
    groupcounter = groupcounter + 1

print("Groups in the file: ")




for x in marketList:
    print(dataList[x])
print('which one do u want to export?')
target = input()
targetStart = 0
targetEnd = 0
target = target + '\n'

if target in dataList:
    targetID = dataList.index(target)
    targetID = marketList.index(targetID)
    targetStart = marketList[targetID]
    targetEnd = marketList[targetID+1]
else:
    print("ERROR")


print("Enter the filename for IKTE Export:")
filerec = input()
filerec = filerec + ".txt"

f = open(filerec, "w+")
for i in range(int(targetStart), int(targetEnd)):
        print(dataList[i])
        f.write("%s\n" % dataList[i])
f.close()
print("Finish. Group is exported")
os.startfile(filerec)