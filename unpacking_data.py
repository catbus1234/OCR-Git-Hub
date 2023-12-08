from zipfile import ZipFile
with ZipFile('data2.zip', 'r') as f:
     f.extractall()