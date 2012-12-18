from setuptools import setup, find_packages

version = '0.0.2'

setup(name='docx-mailmerge',
      version=version,
      description='Performs a Mail Merge on docx (Microsoft Office Word) files',
      long_description=open('README.rst').read(),
      classifiers=[
          'Topic :: Text Processing',
      ],
      author='Bouke Haarsma',
      author_email='bouke@webatoom.nl',
      url='http://github.com/Bouke/docx-mailmerge',
      license='MIT',
      py_modules=['mailmerge'],
      zip_safe=False,
)
