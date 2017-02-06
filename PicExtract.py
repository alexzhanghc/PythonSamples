#!/usr/bin/env python
# -*- coding: utf-8 -*-

import shutil
import os
import zipfile

input_name = input("Enter the title of your file: ")
file_name = input_name + '.docx'
home = os.path.expanduser('~')
desktop = home + '/Desktop/'
path = os.path.join(desktop, file_name)
assert os.path.exists(path), "I did not find the file at " + str(path)

z = zipfile.ZipFile(path)

all_files = z.namelist()

# get all files in word/media/ directory
images = list(filter(lambda x: x.startswith('word/media/'), all_files))

image_folder = '/images/'

if not os.path.exists(desktop + image_folder):
    os.mkdir(desktop + image_folder)

# Extract file
for image in images:
    image_file = z.open(image).read()
    image_name = image.split(r'/').pop()
    f = open(desktop + '/images/'+image_name, 'wb')
    f.write(image_file)

print('Extraction finished succesfully!')
