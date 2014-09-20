===============
docx Mail Merge
===============

.. image:: https://travis-ci.org/Bouke/docx-mailmerge.png?branch=master
    :alt: Build Status
    :target: https://travis-ci.org/Bouke/docx-mailmerge

Performs a Mail Merge on Office Open XML (docx) files. Can be used on any
system without having to install Microsoft Office Word. Supports Python 2.7,
3.2 and up.

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
    document = MailMerge('input.docx')


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
                         {'col1': 'Row 3, Column 1', 'col2': 'Row 3 Column 1'}]


Write document to file. This should be a new file, as ``ZipFile`` cannot modify
existing zip files.
::

    document.write('output.docx')


Unit tests
==========

In order to make sure that the library performs the way it was designed, unit
tests are used. When providing new features, or fixing bugs, there should be a
unit test that demonstrates it. Run the test suite::

    python -m unittest discover


Todo / Wish List
================

* Preserve formatting of the merge field, currently it defaults to the
  formatting of the containing text.
* Image merging.

Contributing
============

* Fork the repository on GitHub and start hacking
* Send a pull request with your changes
