from setuptools import setup
from codecs import open
from os import path

here = path.abspath(path.dirname(__file__))

# Get the long description from the README file
with open(path.join(here, 'README.rst'), encoding='utf-8') as f:
    long_description = f.read()
    
#def readme():
#    with open('README.rst', encoding='utf-8') as f:
#        return f.read()
        
#def licensefile():
#    with open('LICENSE.txt') as f:
#        return f.read()
        
setup(name='pyeviews',
      version='1.0.2',
      description='Data import/export and EViews function calls from Python',
      long_description=long_description,
      long_description_content_type='text/x-rst',
      classifiers=['Development Status :: 5 - Production/Stable',
                   'License :: OSI Approved :: GNU General Public License v3 (GPLv3)',
                   'Topic :: Scientific/Engineering :: Information Analysis',
                   'Programming Language :: Python :: 2',
                   'Programming Language :: Python :: 3',
                   'Operating System :: Microsoft :: Windows'],
      keywords='eviews econometrics',
      url='https://github.com/bexer/pyeviews',
      author='IHS EViews',
      author_email='pyeviews@yahoo.com',
      license='GPLv3',
      packages=['pyeviews'],
      install_requires=['comtypes','numpy','pandas'],
      include_package_data=True,
      zip_safe=False)
