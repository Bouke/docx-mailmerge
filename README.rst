===============
docx Mail Merge
===============

Installation with ``pip``:
::

    $ pip install docx-mailmerge


Usage
=====

Read the file.
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

Write document to file. This should be a new file, as ``ZipFile`` cannot modify
existing zip files.
::

    document.write('output.docx')


Todo / Wish List
================

* Make it easier to work with repeating blocks. This is currently supported,
  but somewhat cumbersome to work with.
* Preserve formatting of the merge field, currently it defaults to the
  formatting of the containing text.
* Image merging.

Contributing
============

* Fork the repository on GitHub and start hacking
* Send a pull request with your changes
