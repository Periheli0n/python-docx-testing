# encoding: utf-8

"""
Custom element classes related to hyperlinks (CT_Hyper).
"""

from ..ns import qn
from ..simpletypes import ST_RelationshipId
from ..xmlchemy import (
    BaseOxmlElement, RequiredAttribute, OneAndOnlyOne
)

class CT_Hyper(BaseOxmlElement):
    """
    ``<w:hyperlink>`` element, containing the properties and text for a hyperlink.
    """
    r = OneAndOnlyOne('w:r')
    rid = RequiredAttribute('r:rId',ST_RelationshipId)

    def _insert_r(self, r):
        if self.r is not None:
            return self.r
        self.insert(0, r) 
        return r

    def clear_content(self):
        """
        Remove all child elements except the ``<w:rPr>`` element if present.
        """
         
        if self.r is not None:
            self.r.clear_content()
        self.remove(self.r)
        rid = None

    @property
    def style(self):
        """
        String contained in w:val attribute of <w:rStyle> grandchild, or
        |None| if that element is not present.
        """
        r = self.r
        if r is None:
            return None
        return r.style

    @style.setter
    def style(self, style):
        """
        Set the character style of this <w:r> element to *style*. If *style*
        is None, remove the style element.
        """
        rPr = self.r.get_or_add_rPr()
        rPr.style = style

    @property
    def text(self):
        """
        A string representing the textual content of this run, with content
        child elements like ``<w:tab/>`` translated to their Python
        equivalent.
        """
        return self.r.text

    @text.setter
    def text(self, text):
        self.r= text
        return self.r.text
