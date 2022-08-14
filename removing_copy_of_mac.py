import os

dir_path = r''
seperator = os.path.sep

files = os.listdir(dir_path)
for file in files:
	old_path = dir_path+ seperator +file
	if "Copy of " in file:
	# if "Travel " in file:
		try:
			# new_name = file.split("Travel ")[1]
			# new_name = "14 Day Prior " + new_name

			new_name = file.split("Copy of ")[1]

			new_path = dir_path+ seperator + new_name
			os.rename(old_path,new_path)
		except:
			print(file + " failed to rename")