# create-certificate

for create certificate you need an powerpoint template (.pptx) and one .csv file which every row of this file consists of all informataions you want to add to template.

the csv file should have a colunm "name" which is the name of output of every row.

# instalation
first you should install *libreoffice4.4* use this [link](https://downloadarchive.documentfoundation.org/libreoffice/old/4.4.7.2/deb/x86_64/) to download and then use [this](https://sourcedigit.com/14874-install-libreoffice-4-4-ubuntu-14-04/) link for install struction.

second install unoconv:
```
pip install unoconv
```

install imagemagick
```commandline
sido apt install imagemagick
```

## Possible issues
During conversion, two errors occur after running the convert command:

```
convert-im6.q16: not authorized `multiple_img.pdf' @ error/constitute.c/ReadImage/412.
convert-im6.q16: no images defined `output-%3d.jpg' @ error/convert.c/ConvertImageCommand/3258.
```

For the first error, you can edit ```/etc/ImageMagick-6/policy.xml``` and change the following line:

```commandline
<policy domain="coder" rights="none" pattern="PDF" />
```

to 
```commandline
<policy domain="coder" rights="read|write" pattern="PDF" />
```

For the second error, this is because ghostscript has not been installed on the system. Try to install it:

```commandline
apt install ghostscript

```

# how to use:
```
python main.py -l path/to/file.csv -t "path/to/file.pptx"
```
