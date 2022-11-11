# create-certificate

for create certificate you need an poerpoint template (.pptx) and one .csv file which every row of this file consists of all informataions you want to add to template.

the csv file should have a colunm "name" which is the name of output of every row.

# instalation
first you should install *libreoffice4.4* use this [link](https://downloadarchive.documentfoundation.org/libreoffice/old/4.4.7.2/deb/x86_64/) to download and then use [this](https://sourcedigit.com/14874-install-libreoffice-4-4-ubuntu-14-04/) link for install struction.

second install unoconv:
```
pip install unoconv
```

# how to use:
```
python main.py -l path/to/file.csv -t "path/to/file.pptx"
```
