===============
docx Mail Merge
===============

.. image:: https://badge.fury.io/py/docx-mailmerge2.png
    :alt: PyPI
    :target: https://pypi.python.org/pypi/docx-mailmerge2

Performs a Mail Merge on Office Open XML (docx) files. Can be used on any
system without having to install Microsoft Office Word. Supports Python 2.7,
3.3 and up.

Installation
============

Installation with ``pip``:
::

    $ pip install docx-mailmerge2


Usage
=====

Open the file.
::

    from mailmerge import MailMerge
    with MailMerge('input.docx',
            remove_empty_tables=False,
            auto_update_fields_on_open="no") as document:
        ...


List all merge fields.
::

    print(document.get_merge_fields())


Merge fields, supplied as kwargs.
::

    document.merge(field1='docx Mail Merge',
                   field2='Can be used for merging docx documents')

Merge table rows. In your template, add a MergeField to the row you would like
to designate as template. Supply the name of this MergeField as ``anchor``
parameter. The second parameter contains the rows with key-value pairs for
the MergeField replacements.

If the tables are empty and you want them removed, set remove_empty_tables=True
in constructor.
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
and separates them by page or section breaks (see function documentation).

When using this feature, make sure you don't use comments, footnotes, 
bookmarks, etc. This is because these elements have an id attribute, which
must be unique. This library does not handle this, resulting in invalid
documents.
::

    document.merge_templates([
        {'field1': "Foo", 'field2: "Copy #1"},
        {'field1': "Bar", 'field2: "Copy #2"},
    ], separator='page_break')


Starting in version 0.6.0 the fields formatting is respected when compatible.
Numeric, Text, Conditional and Date fields (0.6.2) are implemented.
For Date fields a compatible datetime, date or time objects must be provided.
If locale support is needed, make sure to call the setlocale before merging
::

    import locale
    locale.setlocale(locale.LC_TIME, '') # for system locale

    document.merge_templates([
        {'datefield': datetime.date('2022-04-15')},
    ], separator='page_break')

The {NEXT} fields are supported (0.6.3).

You can also use the merge fields inside other fields, for example to insert
pictures in the docx {INCLUDEPICTURE} or for conditional texts {IF}
::

    { INCLUDEPICTURE "{ MERGEFIELD path }/{ MERGEFIELD image }" }
    { IF "{ MERGEFIELD reason }" <> "" "Reason: { MERGEFIELD reason }" }

Always enclose the fields with double quotes, as the MERGE fields will be first
filled in with data and then the other fields will be computed through Word.

If the fields are nested inside other fields, the outer fields need to be
updated in Word. This can be done by selecting everything (CTRL-a) and then
update the fields (F9). There is a way to force the Word to update fields
automatically when opening the document. docx-mailmerge can set this
setting when saving the document. You can configure this feature by using
the *auto_update_fields_on_open* parameter. The value *always* will set the
setting regardless if needed or not and the value *auto* will only set it
when necessary (when nested fields exist). The default value *no* will not
activate this setting.


Write document to file. This should be a new file, as ``ZipFile`` cannot modify
existing zip files.
::

    document.write('output.docx')

See also the unit tests and this nice write-up `Populating MS Word Templates
with Python`_ on Practical Business Python for more information and examples.

Todo / Wish List
================

* Update fields instead of replacing them, so future merges will also work as an "update"
* Create single word documents for each row of data
* Implement SKIPIF and NEXTIF fields

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

| This library was `originally`_ written by `Bouke Haarsma`_ and contributors.
| This repository is maintained by `Iulian Ciorăscu`_.

.. _Bouke Haarsma: https://twitter.com/BoukeHaarsma
.. _Populating MS Word Templates with Python: http://pbpython.com/python-word-template.html
.. _originally: https://github.com/Bouke/docx-mailmerge
.. _Iulian Ciorăscu: https://github.com/iulica/
