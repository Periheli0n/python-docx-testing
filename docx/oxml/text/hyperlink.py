# encoding: utf-8

"""
Custom element classes related to hyperlinks (CT_Hyper).
"""
from .. import parse_xml, OxmlElement
from ..ns import qn
from ..simpletypes import ST_RelationshipId
from ..xmlchemy import (
    BaseOxmlElement, RequiredAttribute, OneAndOnlyOne, ZeroOrOne
)

class CT_Hyper(BaseOxmlElement):
    """
    ``<w:hyperlink>`` element, containing the properties and text for a hyperlink.
    """
    r = ZeroOrOne('w:r')
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

    def set_hyperlink(self,text,rId):
        self.set(qn('r:id'),rId)
        self.add_r()
        self.r.style='Hyperlink'
        self.r.text=text


    @classmethod
    def new_hyperlink(cls,text,rId):
        """
        Return new `w:hyperlink` element that displays *text* and links to the url in relationship *rId*.
        """

        #return parse_xml(cls._hyperlink_xml(text,rId))
        return cls._hyperlink_xml(text,rId)

    @classmethod
    def _hyperlink_xml(cls,text,rId):
        hyp= OxmlElement('w:hyperlink')
        hyp.set(qn('r:id'),rId)
        run= OxmlElement('w:r')
        
        style= OxmlElement('w:rStyle')
        style.set(qn('w:val'),'Hyperlink')

        rPr= OxmlElement('w:rPr')
        rPr.append(style)
        run.append(rPr)
        run.text=text

        hyp.append(run)

        return hyp

        '''
        return (
            '<w:hyperlink r:id="%s">\n'
            '    <w:r>\n'
            '        <w:rPr>\n'
            '            <w:rStyle w:val="Hyperlink"/>\n'
            '        </w:rPr>\n'
            '        <w:t>%s</w:t>\n'
            '    </w:r>\n'
            '</w:hyperlink>\n'
        ) % (
            rId,
            text
        )
        '''

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
