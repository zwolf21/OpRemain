#-*-coding:utf-8
#setup.py
from distutils.core import setup
import py2exe, sys

sys.argv.append("py2exe")

setup(
	windows=[{'script':"OpRemain.py"}],
	options={
	    "py2exe":{
	        "packages" : [], "bundle_files":1,
	        "optimize":2, 
	        } 
	    }, 
    zipfile = None 
)