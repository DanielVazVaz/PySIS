# PySIS

Abstraction layer over the com HYSYS interface using Python. Allows for the use of functions without having to step to com level, which is sometimes esoteric. Maintains the whole functionality of the COM objects, since these are loaded as attributes of the superclass.

As of now, it is checked to work with `Aspen HYSYS V11`. No idea if it works with `Aspen HYSYS V12`.

## win32 DLL problem
Right now, it looks like for Python 3.8. there are problems with the `win32api` package. This worked to me:
```
pip install pywin32==225
```

If this still does not work, make sure that you do not have other `pywin32` in your environment, e.g., some version installed with `conda`. 