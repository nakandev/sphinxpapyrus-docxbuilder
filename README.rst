=========================
sphinxpapyrus-docxbuilder
=========================

sphinxpapyrus-docxbuilder is a Sphinx extension for Word (.docx) file output.

Features
========

* Docx file as style template
* Inline Markup
* Headings
* Bullet / Enumerated Multilevel List
* Definition List
* Field List / Option List as 2 collumn table
* Blocks
* Simple Table / Grid Table (surpport sppaning, nesting)
* Transitions
* Image / Figure
* Footnotes as normal paragraph

Requirements
------------

* Sphinx>=1.3
* python-docx==0.8.6

Installation
------------

Run the following command::

   pip install sphinxpapyrus-docxbuilder

Usage
-----

Add the extension module name into *conf.py* in your Sphinx document::

   extentions = ['sphinxpapyrus.docxbuilder']

Optionally, you can set style file::

   docx_style = 'mystyle.docx'

You can also set docx core properties::

   docx_coreproperties = {
       'title': 'Jelly Island Murders',
       'author': 'Arashiyama Hotori',
   }

For more properties, see `python-docx ducument`__ .

Other docx options::

   # Grouping the document tree into Docx files. List of tuples
   # (source start file, target name, {coreproperties}).
   docx_documents = [
       (master_doc, project, {
           'title': 'Document Title',
           'author': 'Author',
       }),
   ]
   docx_pagebreak_level = 2  # insert page break before each heading 1, 2 and title
   docx_imagetable_align = 'center'  # 'left', 'center', or 'right'

__ https://python-docx.readthedocs.io/en/latest/api/document.html#docx.opc.coreprops.CoreProperties

Finaly, output docx with following command::

   make docx
