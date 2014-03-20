__author__ = "TKretts"
__date__ = "$20.03.2014 08:49:47$"

from setuptools import setup, find_packages

setup(
    name='pyssrs',
    version='1.0',
    packages=find_packages(),

    install_requires=[
        'requests',
        'lxml',
    ],

    author='TKretts',
    author_email='tkretts666@gmail.com',

    url='',
    license='',
    long_description='Python-module for SQl Server Reporting Services',

)