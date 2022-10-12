# PySIS

[![Documentation Status](https://readthedocs.org/projects/pysis-doc/badge/?version=latest)](https://pysis-doc.readthedocs.io/en/latest/?badge=latest)

Abstraction layer over the COM HYSYS interface using Python. Allows for the use of functions without having to step to COM level, which is sometimes esoteric. Maintains the whole functionality of the COM objects, since these are loaded as attributes of the superclass.

As of now, it is checked to work with `Aspen HYSYS V11`. No idea if it works with `Aspen HYSYS V12`.

## Installation

Install the latest version of this repository to your machine following one of the options below accordingly to your preferences:

- users with git:<br/>
<pre>git clone https://github.com/DanielVazVaz/PySIS.git
cd PySIS
pip install -e .
</pre>

- users without git:<br/>
Browser to https://github.com/DanielVazVaz/PySIS, click on the `Code` button and select `Download ZIP`. Unzip the files from your Download folder to the desired one. Open a terminal inside the folder you just unzipped (make sure this is the folder containing the `setup.py` file). Run the following command in the terminal:
<pre>
pip install -e .
</pre>

- contributors:<br/>
<pre>git clone https://github.com/DanielVazVaz/PySIS.git
cd PySIS
pip install -e .[dev]
</pre>

## win32 DLL problem

Right now, it looks like for Python 3.8. there are problems with the `win32api` package. This worked for me:

```
pip install pywin32==225
```

If this still does not work, make sure that you do not have other `pywin32` in your environment, e.g., some version installed with `conda`.
