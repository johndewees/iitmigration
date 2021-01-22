filename = 'test.txt'
filehandle = open('test.txt')

row = 4608
filelist = list()

for item in filehandle:
    print(item) 
    if item.startswith('4'):
        item = item.strip()
        item = int(item)     
        if item == row:
            filelist.append(item)
            row = row + 1
        else:
            print(filelist)
            break