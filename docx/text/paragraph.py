# encoding: utf-8

"""
Paragraph-related proxy types.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from ..enum.style import WD_STYLE_TYPE
from .parfmt import ParagraphFormat
from .run import Run
from ..shared import Parented
from .hyperlink import Hyperlink
from ..oxml.text.hyperlink import CT_Hyper
from ..oxml import OxmlElement
from ..oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE


class Paragraph(Parented):
    """
    Proxy object wrapping ``<w:p>`` element.
    """
    def __init__(self, p, parent):
        super(Paragraph, self).__init__(parent)
        self._p = self._element = p
        self._bookmark = ""


    def add_run(self, text=None, style=None):
        """
        Append a run to this paragraph containing *text* and having character
        style identified by style ID *style*. *text* can contain tab
        (``\\t``) characters, which are converted to the appropriate XML form
        for a tab. *text* can also include newline (``\\n``) or carriage
        return (``\\r``) characters, each of which is converted to a line
        break.
        """
        r = self._p.add_r()
        run = Run(r, self)
        if text:
            run.text = text
        if style:
            run.style = style
        return run

    '''
    def add_hyperlink(self, url, text):
        rId = self.part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

        #hypr_elm = self._p.add_hypr({'rId': rId})
        hypr_elm= self._p.add_hypr(attrs=[('rId',rId)])
        hypr_obj = Hyperlink(hypr_elm, rId, self)
        hypr_obj.add_text(text)

        return hypr_obj
    '''
    def add_hyperlink(self, text, url):
        rId = self.part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
        hyp = CT_Hyper.new_hyperlink(text,rId)
        self._p.append(hyp)
        return hyp

    def add_hyperlink2(self,text,url):
        rId = self.part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
        hyp = self._p.add_hypr()
        hyp.set_hyperlink(text,rId)
        return hyp

    def create_bookmark(self, id, name):
        if not (self.style.name == 'Heading 1' or self.style.name == 'Heading 2'):
            raise ValueError(f"Bookmarked headers can only be created for paragraphs of style 'Heading 1' or 'Heading 2'. Paragraph style: '{self.style.name}'")
        
        start= OxmlElement('w:bookmarkStart')
        start.set(qn('w:id'),str(id))
        start.set(qn('w:name'),name)
        end= OxmlElement('w:bookmarkEnd')
        end.set(qn('w:id'),str(id))
        self._p.insert(1, start)
        self._p.append(end)
        self.bookmark= name

    @property
    def alignment(self):
        """
        A member of the :ref:`WdParagraphAlignment` enumeration specifying
        the justification setting for this paragraph. A value of |None|
        indicates the paragraph has no directly-applied alignment value and
        will inherit its alignment value from its style hierarchy. Assigning
        |None| to this property removes any directly-applied alignment value.
        """
        return self._p.alignment

    @alignment.setter
    def alignment(self, value):
        self._p.alignment = value

    def clear(self):
        """
        Return this same paragraph after removing all its content.
        Paragraph-level formatting, such as style, is preserved.
        """
        self._p.clear_content()
        return self

    def insert_paragraph_before(self, text=None, style=None):
        """
        Return a newly created paragraph, inserted directly before this
        paragraph. If *text* is supplied, the new paragraph contains that
        text in a single run. If *style* is provided, that style is assigned
        to the new paragraph.
        """
        paragraph = self._insert_paragraph_before()
        if text:
            paragraph.add_run(text)
        if style is not None:
            paragraph.style = style
        return paragraph

    @property
    def paragraph_format(self):
        """
        The |ParagraphFormat| object providing access to the formatting
        properties for this paragraph, such as line spacing and indentation.
        """
        return ParagraphFormat(self._element)

    @property
    def runs(self):
        """
        Sequence of |Run| instances corresponding to the <w:r> elements in
        this paragraph.
        """
        return [Run(r, self) for r in self._p.r_lst]

    @property
    def style(self):
        """
        Read/Write. |_ParagraphStyle| object representing the style assigned
        to this paragraph. If no explicit style is assigned to this
        paragraph, its value is the default paragraph style for the document.
        A paragraph style name can be assigned in lieu of a paragraph style
        object. Assigning |None| removes any applied style, making its
        effective value the default paragraph style for the document.
        """
        style_id = self._p.style
        return self.part.get_style(style_id, WD_STYLE_TYPE.PARAGRAPH)

    @style.setter
    def style(self, style_or_name):
        style_id = self.part.get_style_id(
            style_or_name, WD_STYLE_TYPE.PARAGRAPH
        )
        self._p.style = style_id

    @property
    def text(self):
        """
        String formed by concatenating the text of each run in the paragraph.
        Tabs and line breaks in the XML are mapped to ``\\t`` and ``\\n``
        characters respectively.

        Assigning text to this property causes all existing paragraph content
        to be replaced with a single run containing the assigned text.
        A ``\\t`` character in the text is mapped to a ``<w:tab/>`` element
        and each ``\\n`` or ``\\r`` character is mapped to a line break.
        Paragraph-level formatting, such as style, is preserved. All
        run-level formatting, such as bold or italic, is removed.
        """
        text = ''
        for run in self.runs:
            text += run.text
        return text

    @text.setter
    def text(self, text):
        self.clear()
        self.add_run(text)

    def _insert_paragraph_before(self):
        """
        Return a newly created paragraph, inserted directly before this
        paragraph.
        """
        p = self._p.add_p_before()
        return Paragraph(p, self._parent)

    @property 
    def bookmark(self):
        """
        String used by Table of Contents (Heading 1 & Heading 2 specifically) to reference these headers.
        """
        return self._bookmark

    @bookmark.setter
    def bookmark(self,bookmark):
        self._bookmark = bookmark