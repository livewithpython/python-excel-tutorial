Python Excel Tutorial
=====================

How to install
===============
1. Install python
2. pip install ipython pyzmq jinja2 tornado pandas xlrd openpyxl xlwt

if it's not working :
	pip install matplotlib
else :
	pip install -e git+https://github.com/matplotlib/matplotlib.git#egg=matplotlib


Step 5: matplotlib
The first of these tools we need is pkg-config which helps to configure matplotlib when it is being compiled.

$ brew install pkg-config

Next we need to install the libpng library as well as freetype as matplotlib will need to access these libraries in order to compile

$ brew install freetype

$ brew install libpng

$ brew install ffmpeg

We've added ffmpeg to allow us to save movies while using the excellent matplotlib.animation library. Now we are ready to install matplotlib
$ pip install matplotlib


Ubuntu :

I had this issue on Ubuntu server 12.04.

I had to install libfreetype6-dev and libpng-dev from the repositories. I was using a virtualenv and installing matplotlib using pip when I ran into this issue.

Hints that I needed to do this came from the warning messages that popup early in the matplotlib installation so keep an eye out for those messages which indicate a dependency is found, but not the headers.

brew install gfortran

6. pip install scikit-learn scipy
7.sudo chmod a+w /usr/local/Cellar

8. git clone git@github.com:livewithpython/python-excel-tutorial.git
   or download from https://github.com/livewithpython/python-excel-tutorial

Rubayeet Vai, the following links can help you : 

1. http://www.daniweb.com/software-development/python/threads/186575/reading-excel-file-in-python
2. http://www.simplistix.co.uk/presentations/python-excel.pdf
3. https://github.com/python-excel/xlwt/tree/master/xlwt/examples
4. http://www.python-excel.org/
5. https://github.com/ipython/ipython/wiki/A-gallery-of-interesting-IPython-Notebooks
6. http://www.tapir.caltech.edu/~dtsang/python.html
7. %pylab inline
