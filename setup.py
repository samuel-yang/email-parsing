#setup.py

from distutils.core import setup
import py2exe

setup(console=['Source_Compiler.py'],
      options= {
          'py2exe': {
              'packages': ['apiclient', 'oauth2client', 'base64', 'googleapiclient', 'httplib2', 'os', 'io', 'xlrd', 'datetime', 're']
          }
      }
)

