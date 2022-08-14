from os import walk
import os

mypath = ""
new_path = ""

f = []
for (dirpath, dirnames, filenames) in walk(mypath):
    f.extend(filenames)
    break

for old_name in f:
    os.rename(mypath + old_name, new_path + old_name)
