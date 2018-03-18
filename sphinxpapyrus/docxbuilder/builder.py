# -*- coding: utf-8 -*-
"""
    sphinxpapyrus.docxbuilder.builder
    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    docx Sphinx builder.

    :copyright: Copyright 2018 by nakandev.
    :license: MIT, see LICENSE for details.
"""

import codecs
from os import path

from docutils import nodes
from docutils.io import StringOutput

from sphinx.builders import Builder
from sphinx.util import logging
from sphinx.util.osutil import ensuredir, os_path
from sphinx.util.console import bold, darkgreen, brown
from .writer import DocxWriter, DocxTranslator

if False:
    # For type annotation
    from typing import Any, Dict, Iterator, Set  # NOQA
    from docutils import nodes  # NOQA
    from sphinx.application import Sphinx  # NOQA

logger = logging.getLogger(__name__)

def inline_all_toctrees(builder, docnameset, docname, tree, colorfunc, traversed):
    # type: (Builder, Set[unicode], unicode, nodes.Node, Callable, nodes.Node) -> nodes.Node
    """Inline all toctrees in the *tree*.

    Record all docnames in *docnameset*, and output docnames with *colorfunc*.
    """
    from six import text_type
    from docutils import nodes
    from sphinx import addnodes

    tree = tree.deepcopy()
    for toctreenode in tree.traverse(addnodes.toctree):
        newnodes = []
        includefiles = map(text_type, toctreenode['includefiles'])
        for includefile in includefiles:
            if includefile not in traversed:
                try:
                    traversed.append(includefile)
                    logger.info(colorfunc(includefile) + " ", nonl=1)
                    subtree = inline_all_toctrees(builder, docnameset, includefile,
                                                  builder.env.get_doctree(includefile),
                                                  colorfunc, traversed)
                    docnameset.add(includefile)
                except Exception:
                    logger.warning('toctree contains ref to nonexisting file %r',
                                   includefile, location=docname)
                else:
                    sof = addnodes.start_of_file(docname=includefile)
                    sof.children = subtree.children
                    for sectionnode in sof.traverse(nodes.section):
                        if 'docname' not in sectionnode:
                            sectionnode['docname'] = includefile
                    newnodes.append(sof)
        toctreenode.parent['numbered'] = toctreenode['numbered']
        toctreenode.parent.replace(toctreenode, newnodes)
    return tree

class DocxBuilder(Builder):
    name = 'docx'
    format = 'docx'
    out_suffix = '.docx'
    allow_parallel = False
    default_translator_class = DocxTranslator

    current_docname = None  # type: unicode

    def init(self):
        # type: () -> None
        pass

    def get_outdated_docs(self):
        # type: () -> Iterator[unicode]
        return 'all documents'

    def get_target_uri(self, docname, typ=None):
        # type: (unicode, unicode) -> unicode
        return ''

    def fix_refuris(self, tree):
        # type: (nodes.Node) -> None
        # fix refuris with double anchor
        fname = self.config.master_doc + self.out_suffix
        for refnode in tree.traverse(nodes.reference):
            if 'refuri' not in refnode:
                continue
            refuri = refnode['refuri']
            hashindex = refuri.find('#')
            if hashindex < 0:
                continue
            hashindex = refuri.find('#', hashindex + 1)
            if hashindex >= 0:
                refnode['refuri'] = fname + refuri[hashindex:]

    def prepare_writing(self, docnames):
        # type: (Set[unicode]) -> None
        self.writer = DocxWriter(self)

    def assemble_doctree(self, start=None):
        # type: () -> nodes.Node
        master = start if start else self.config.master_doc
        tree = self.env.get_doctree(master)
        tree = inline_all_toctrees(self, set(), master, tree, darkgreen, [master])
        tree['docname'] = master
        self.env.resolve_references(tree, master, self)
        self.fix_refuris(tree)
        return tree

    def assemble_toc_fignumbers(self):
        new_fignumbers = {}  # type: Dict[unicode, Dict[unicode, Tuple[int, ...]]]
        # {u'foo': {'figure': {'id2': (2,), 'id1': (1,)}}, u'bar': {'figure': {'id1': (3,)}}}
        for docname, fignumlist in self.env.toc_fignumbers.items():
            for figtype, fignums in fignumlist.items():
                alias = "%s/%s" % (docname, figtype)
                new_fignumbers.setdefault(alias, {})
                for id, fignum in fignums.items():
                    new_fignumbers[alias][id] = fignum

        return {self.config.master_doc: new_fignumbers}

    def write(self, *ignored):
        # type: (Any) -> None
        docnames = self.env.all_docs
        if self.config.docx_documents:
            docx_documents = self.config.docx_documents
        else:
            docx_documents = [(self.config.master_doc, self.config.project,
                              self.config.docx_coreproperties)]
        for entry in docx_documents:
            start, name, coreproperties = entry
            self.config.docx_coreproperties = coreproperties
            logger.info(bold('preparing documents... '), nonl=True)
            self.prepare_writing(docnames)
            logger.info('done')

            logger.info(bold('assembling single document... '), nonl=True)
            doctree = self.assemble_doctree(start)
            self.env.toc_fignumbers = self.assemble_toc_fignumbers()
            logger.info('')
            logger.info(bold('writing... '), nonl=True)
            docname = [start, name]
            self.write_doc(docname, doctree)
            logger.info('done')

    def write_doc(self, docname, doctree):
        # type: (unicode, nodes.Node) -> None
        start, name = docname
        self.current_docname = start
        self.fignumbers = self.env.toc_fignumbers.get(start, {})
        destination = StringOutput(encoding='utf-8')
        self.writer.write(doctree, destination)
        outfilename = path.join(self.outdir, os_path(name) + self.out_suffix)
        ensuredir(path.dirname(outfilename))
        try:
            self.writer.save(outfilename)
        except (IOError, OSError) as err:
            logger.warning("error writing file %s: %s", outfilename, err)

    def finish(self):
        # type: () -> None
        pass

