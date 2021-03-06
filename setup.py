import os.path
from setuptools import setup

# Get a handy base dir
project_dir = os.path.abspath(os.path.dirname(__file__))

with open(os.path.join(project_dir, 'README.rst')) as f:
    README = f.read()

setup(
    name='pyxlsb',
    version='1.1.0',

    description='Excel 2007+ Binary Workbook (xlsb) parser',
    long_description=README,

    author='William Turner',
    author_email='willtur.will@gmail.com',

    url='https://github.com/wwwiiilll/pyxlsb',

    license='MIT',

    classifiers=[
        'Development Status :: 5 - Production/Stable',

        'License :: OSI Approved :: MIT License',

        'Programming Language :: Python :: 2',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6'
        'Programming Language :: Python :: 3.7'
    ],

    packages=['pyxlsb'],

    install_requires=[
        'enum34'
    ],

    zip_safe=False
)
