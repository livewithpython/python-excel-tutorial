Python Excel Tutorial
=====================

How to install
===============

1. pip install ipython pyzmq jinja2 tornado pandas xlrd openpyxl xlwt

2. pip install matplotlib

In case you get exception for matplotlib then following solution may work.

Mac: 
===
The first of these tools we need is pkg-config which helps to configure matplotlib when it is being compiled.

$ brew install pkg-config

Next we need to install the libpng library as well as freetype as matplotlib will need to access these libraries in order to compile

$ brew install freetype

$ brew install libpng

$ brew install ffmpeg

We've added ffmpeg to allow us to save movies while using the excellent matplotlib.animation library. Now we are ready to install matplotlib

$ pip install matplotlib


Ubuntu :
========
Install libfreetype6-dev and libpng-dev from the repositories. 


3. pip install scikit-learn scipy
In case you get error for fortran compiler then you have install gfortran.
$ brew install gfortran
If you get permission denired error for the /usr/local/Cellar directory, run the following command:
$ sudo chmod a+w /usr/local/Cellar

4. git clone git@github.com:livewithpython/python-excel-tutorial.git
   or download from https://github.com/livewithpython/python-excel-tutorial


References:
============
1. http://www.python-excel.org/
2. https://github.com/python-excel/xlwt/tree/master/xlwt/examples
3. http://www.simplistix.co.uk/presentations/python-excel.pdf
4. https://github.com/ipython/ipython/wiki/A-gallery-of-interesting-IPython-Notebooks
5. http://www.tapir.caltech.edu/~dtsang/python.html
