# encoding: utf-8

"""
Run-related proxy objects for python-docx, Run in particular.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..enum.style import WD_STYLE_TYPE
from ..enum.text import WD_BREAK
from .font import Font
from ..shape import InlineShape
from ..shared import Parented
from ..oxml.text.hyperlink import CT_Hyper
from .run import _Text

class Hyperlink(Parented):
    """
    Proxy object wrapping ``<w:hyperlink>`` element. Hyperlink has a required property of a Relationship ID corresponding to a document rels proprty with the URL. Text attributes are determined by an internal Run nested inside the hyperlink.
    """
    def __init__(self, hyper, rId, parent):
        super(Hyperlink, self).__init__(parent)
        #self._hyper = self._element = self.element = CT_Hyper.new_hyperlink(text,rId)
        self._hyper = self._element = self.element = hyper
        self.element.add_r()
        



    def add_text(self, text):
        """
        Sets text of the inner run element.
        """
        t = self.element.r.add_t(text)
        return _Text(t)

    def clear(self):
        """
        Return reference to this run after removing all its content. All run
        formatting is preserved.
        """
        self.element.r.clear_content()
        return self

    @property
    def font(self):
        """
        The |Font| object providing access to the character formatting
        properties for this run, such as font name and size.
        """
        return Font(self._element.r)

    @property
    def italic(self):
        """
        Read/write tri-state value. When |True|, causes the text of the run
        to appear in italics.
        """
        return self.element.r.font.italic

    @italic.setter
    def italic(self, value):
        self.element.r.font.italic = value

    @property
    def style(self):
        """
        Read/write. A |_CharacterStyle| object representing the character
        style applied to this run. The default character style for the
        document (often `Default Character Font`) is returned if the run has
        no directly-applied character style. Setting this property to |None|
        removes any directly-applied character style.
        """
        style_id = self.element.r.style
        return self.part.get_style(style_id, WD_STYLE_TYPE.CHARACTER)

    @style.setter
    def style(self, style_or_name):
        style_id = self.part.get_style_id(
            style_or_name, WD_STYLE_TYPE.CHARACTER
        )
        self.element.r.style = style_id

    @property
    def text(self):
        """
        String formed by concatenating the text equivalent of each run
        content child element into a Python string. Each ``<w:t>`` element
        adds the text characters it contains. A ``<w:tab/>`` element adds
        a ``\\t`` character. A ``<w:cr/>`` or ``<w:br>`` element each add
        a ``\\n`` character. Note that a ``<w:br>`` element can indicate
        a page break or column break as well as a line break. All ``<w:br>``
        elements translate to a single ``\\n`` character regardless of their
        type. All other content child elements, such as ``<w:drawing>``, are
        ignored.

        Assigning text to this property has the reverse effect, translating
        each ``\\t`` character to a ``<w:tab/>`` element and each ``\\n`` or
        ``\\r`` character to a ``<w:cr/>`` element. Any existing run content
        is replaced. Run formatting is preserved.
        """
        return self.element.r.text

    @text.setter
    def text(self, text):
        self.element.r.text = text

    @property
    def underline(self):
        """
        The underline style for this |Run|, one of |None|, |True|, |False|,
        or a value from :ref:`WdUnderline`. A value of |None| indicates the
        run has no directly-applied underline value and so will inherit the
        underline value of its containing paragraph. Assigning |None| to this
        property removes any directly-applied underline value. A value of
        |False| indicates a directly-applied setting of no underline,
        overriding any inherited value. A value of |True| indicates single
        underline. The values from :ref:`WdUnderline` are used to specify
        other outline styles such as double, wavy, and dotted.
        """
        return self.element.r.font.underline

    @underline.setter
    def underline(self, value):
        self.element.r.font.underline = value

