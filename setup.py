from setuptools import setup

def readme():
    with open('README.rst') as f:
        return f.read()
        
def licensefile():
    with open('LICENSE.txt') as f:
        return f.read()
        
setup(name='pyeviews',
      version='0.1',
      description='Data import/export and EViews function calls from Python',
      long_description=readme(),
      classifiers=['Development Status :: 4 - Beta',
                   'License :: OSI Approved :: GNU General Public License v3 (GPLv3)',
                   'Topic :: Scientific/Engineering :: Information Analysis'],
      keywords='eviews econometrics',
      url='https://github.com/bexer/pyeviews',
      author='Rebecca Erwin & Steve Yoo',
      author_email='bexer@yahoo.com',
      license=licensefile(),
      packages=['pyeviews'],
      install_requires=['comtypes','numpy','pandas'],
      include_package_data=True,
      zip_safe=False)
