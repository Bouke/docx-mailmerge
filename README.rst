===============
docx Mail Merge
===============

.. image:: https://travis-ci.org/Bouke/docx-mailmerge.png?branch=master
    :alt: Build Status
    :target: https://travis-ci.org/Bouke/docx-mailmerge

.. image:: https://badge.fury.io/py/docx-mailmerge.png
    :alt: PyPI
    :target: https://pypi.python.org/pypi/docx-mailmerge

Performs a Mail Merge on Office Open XML (docx) files. Can be used on any
system without having to install Microsoft Office Word. Supports Python 2.7,
3.3 and up.

Installation
============

Installation with ``pip``:
::

    $ pip install docx-mailmerge


Usage
=====

Open the file.
::

    from mailmerge import MailMerge
    with MailMerge('input.docx') as document:
        ...


List all merge fields.
::

    print document.get_merge_fields()


Merge fields, supplied as kwargs.
::

    document.merge(field1='docx Mail Merge',
                   field2='Can be used for merging docx documents')

Merge table rows. In your template, add a MergeField to the row you would like
to designate as template. Supply the name of this MergeField as ``anchor``
parameter. The second parameter contains the rows with key-value pairs for
the MergeField replacements.
::

    document.merge_rows('col1',
                        [{'col1': 'Row 1, Column 1', 'col2': 'Row 1 Column 1'},
                         {'col1': 'Row 2, Column 1', 'col2': 'Row 2 Column 1'},
                         {'col1': 'Row 3, Column 1', 'col2': 'Row 3 Column 1'}])


Starting in version 0.2.0 you can also combine these two separate calls into a
single call to `merge`.
::

    document.merge(field1='docx Mail Merge',
                   col1=[
                       {'col1': 'A'},
                       {'col1': 'B'},
                   ])


Starting in version 0.2.0 there's also the feature for template merging.
This creates a copy of the template for each item in the list, does a merge,
and separates them by page breaks.
::

    document.merge_pages([
        {'field1': "Foo", 'field2: "Copy #1"},
        {'field1': "Bar", 'field2: "Copy #2"},
    ])


Write document to file. This should be a new file, as ``ZipFile`` cannot modify
existing zip files.
::

    document.write('output.docx')

See also the unit tests and this nice write-up `Populating MS Word Templates
with Python`_ on Practical Business Python for more information and examples.

Todo / Wish List
================

* Image merging.


Contributing
============

* Fork the repository on GitHub and start hacking
* Create / fix the unit tests
* Send a pull request with your changes

Unit tests
----------

In order to make sure that the library performs the way it was designed, unit
tests are used. When providing new features, or fixing bugs, there should be a
unit test that demonstrates it. Run the test suite::

    python -m unittest discover

Credits
=======

This library was written by `Bouke Haarsma`_ and contributors.

.. _Bouke Haarsma: https://twitter.com/BoukeHaarsma
.. _Populating MS Word Templates with Python: http://pbpython.com/python-word-template.html
