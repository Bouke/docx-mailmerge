#!/usr/bin/env python2
from setuptools import setup, find_packages

version = '0.1.1'

setup(name='docx-mailmerge',
      version=version,
      description='Performs a Mail Merge on docx (Microsoft Office Word) files',
      long_description=open('README.rst').read(),
      install_requires = ['lxml>=3.1.2', ],
      classifiers=[
          'License :: OSI Approved :: MIT License',
          'Programming Language :: Python :: 2.7',
          'Programming Language :: Python :: 3.3',
          'Topic :: Text Processing',
      ],
      author='Bouke Haarsma',
      author_email='bouke@webatoom.nl',
      url='http://github.com/Bouke/docx-mailmerge',
      license='MIT',
      py_modules=['mailmerge'],
      zip_safe=False,
      test_suite="tests",
)
