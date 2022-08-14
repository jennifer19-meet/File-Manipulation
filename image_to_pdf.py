from PIL import Image
import os

# The only thing that needs to change is the dir_path (Paste the path inbetween the '')
dir_path = r''

seperator = os.path.sep
files = os.listdir(dir_path)

for file in files:
	if file.lower().endswith(('.png', '.jpg', '.jpeg', '.tiff')):
		try:
			dot_index = file.index(".")
			new_filename = file.strip(".jpeg.jpg.tiff.png") + ".pdf"
			old_path = dir_path + seperator + file
			new_path = dir_path + seperator + new_filename
			new_path_raw = r"{}".format(new_path)

			image = Image.open(old_path)
			rgb_image = image.convert('RGB')
			rgb_image.save(new_path_raw)

			os.remove(old_path)

		except:
			print(file + ' failed to convert to a pdf')
