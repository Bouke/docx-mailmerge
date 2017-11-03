from setuptools import setup

setup(name='docx-mailmerge',
      version='0.4.0',
      description='Performs a Mail Merge on docx (Microsoft Office Word) files',
      long_description=open('README.rst').read(),
      classifiers=[
          'License :: OSI Approved :: MIT License',
          'Programming Language :: Python :: 2.7',
          'Programming Language :: Python :: 3.2',
          'Programming Language :: Python :: 3.3',
          'Programming Language :: Python :: 3.4',
          'Topic :: Text Processing',
      ],
      author='Bouke Haarsma',
      author_email='bouke@webatoom.nl',
      url='http://github.com/Bouke/docx-mailmerge',
      license='MIT',
      py_modules=['mailmerge'],
      zip_safe=False,
      install_requires=['lxml']
)
