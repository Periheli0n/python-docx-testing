# encoding: utf-8

"""
Custom element classes related to table of contents (CT_TOC)
"""

from .ns import qn
from .xmlchemy import BaseOxmlElement, ZeroOrMore, ZeroOrOne, OxmlElement
#from .. import OxmlElement

class CT_SDT(BaseOxmlElement):
    """
    ``<w:sdt>`` element, that is built out into the Table of Contents.
    XML Structure for ToC is complex, but a basic breakdown is as follows:
    <w:sdt>
        <w:sdtPr>
            <w:rPr>
                #Stylings
            </w:rPr>
            <w:id/>
            <w:docPartObj>
			    <w:docPartGallery w:val="Table of Contents"/>
			    <w:docPartUnique/>
		    </w:docPartObj>
        </w:sdtPr>
        <w:sdtEndPr>
            <w:rPr>
                #Stylings
            </w:rPr>
        </w:sdtEndPr>
        <w:sdtContent>
            <w:p>
                #TOC Header/Entries(hyperlinks)
            </w:p>
        </w:sdtContent>
    </w:sdt>
    """
    sdtPr= ZeroOrOne('w:sdtPr')
    sdtEndPr= ZeroOrOne('w:sdtEndPr')
    sdtContent= ZeroOrOne('w:sdtContent') 


    def init_toc(self):
        self._init_sdtPr()
        self._init_sdtEndPr()
        self._init_sdtContent()

    def _init_sdtPr(self):
        """
        This function creates the w:sdtPr element and initializes all settings for ToC. Note that all ToCs created with this function will have an id of '171613375'.
        """
        if self.sdtPr is None:
            self.add_sdtPr()
            #self.sdtPr = OxmlElement('w:sdtPr')
        self.sdtPr.append(self._create_toc_stylings())

        wid= OxmlElement('w:id')
        wid.set(qn('w:val'),'171613375')
        self.sdtPr.append(wid)

        doc_prt_obj= OxmlElement('w:docPartObj')
        doc_prt_gal= OxmlElement('w:docPartGallery')
        doc_prt_gal.set(qn('w:val'),'Table of Contents')
        doc_prt_obj.append(doc_prt_gal)
        doc_prt_obj.append(OxmlElement('w:docPartUnique'))
        self.sdtPr.append(doc_prt_obj)

    def _create_toc_stylings(self):
        """
        This function creates the w:rPr element and initializes all stylings to be added to the ToC.
        """
        stylings= OxmlElement('w:rPr')
        fonts= OxmlElement('w:rFonts')
        fonts.set(qn('w:asciiTheme'),'minorHAnsi')
        fonts.set(qn('w:hAnsiTheme'),'minorHAnsi')
        #fonts.set(qn('w:asciiTheme'),'minorHAnsi')
        #fonts.set(qn('w:asciiTheme'),'minorHAnsi')
        stylings.append(fonts)

        bold= OxmlElement('w:b')
        bold.set(qn('w:val'),'0')
        stylings.append(bold)

        italics= OxmlElement('w:i')
        stylings.append(italics) 

        colors= OxmlElement('w:color')
        colors.set(qn('w:val'), '595959')
        colors.set(qn('w:themeColor'),'text1')
        colors.set(qn('w:themeTint'),'A6')
        stylings.append(colors)

        size= OxmlElement('w:sz')
        size.set(qn('w:val'),'24')
        stylings.append(size)

        return stylings

    def _init_sdtEndPr(self):
        """
        This function creates the w:sdtEndPr element and initializes required child elements.
        """
        if self.sdtEndPr is None:
            self.sdtEndPr= OxmlElement('w:sdtEndPr')
        
        rPr= OxmlElement('w:rPr')
        rPr.append(OxmlElement('w:bCs'))
        italics = OxmlElement('w:i')
        italics.set(qn('w:val'),'0')
        rPr.append(italics)
        rPr.append(OxmlElement('w:noProof'))

        self.sdtEndPr.append(rPr)
    
    def _init_sdtContent(self):
        """
        This function creates the w:sdtContent element and initializes the first w:p element that will always be present.
        """
        if self.sdtContent is None:
            self.sdtContent= OxmlElement('w:sdtContent')

        p= OxmlElement('w:p')
        pPr= OxmlElement('w:pPr')
        pStyle= OxmlElement('w:pStyle')
        pStyle.set(qn('w:val'),'TOCHeading')
        pPr.append(pStyle)
        p.append(pPr)
        t= OxmlElement('w:t')
        t.text="Table of Contents:"
        r=OxmlElement('w:r')
        r.append(t)
        p.append(r)

        self.sdtContent.append(p)
        
    def _add_closing_p(self):
        """
        This function created the final w:p element of the w:sdtContent element. This should be called after doc has been constructed prior to saving.
        """
        p = OxmlElement('w:p')
        pPr = OxmlElement('w:pPr')
        tabs= OxmlElement('w:tabs')
        tab= OxmlElement('w:tab')
        tab.set(qn('w:val'),'left')
        tab.set(qn('w:pos'),'5670')
        tabs.append(tab)
        indt= OxmlElement('w:indt')
        indt.set(qn('w:left'),'5103')
        pPr.append(tabs)
        pPr.append(indt) 
        p.append(pPr)

        run= OxmlElement('w:r')
        rPr= OxmlElement('w:rPr')
        rPr.append(OxmlElement('w:b'))
        rPr.append(OxmlElement('w:bCs'))
        rPr.append(OxmlElement('w:noProof'))
        fieldChar= OxmlElement('w:fldChar')
        fieldChar.set(qn('w:fldCharType'),'end')      
        run.append(rPr)
        run.append(fieldChar)
        p.append(run)

        self.sdtContent.append(p)

    def _create_first_p(self,heading):
        """
        Function used by populate_toc to handle the specific requirements of the first w:p element related to a reference heading in the w:sdtContent element.
        """

    def _create_toc_hyperlink(self,heading):
        """
        Function used to create the w:p elements corresponding to the heading entries in the ToC.
        """

    
    def populate_toc(self, headings):
        first_run= True
        for heading in headings:
            if first_run:
                p= self._create_first_p(heading)
                first_run =False
            else:
                p= self._create_toc_hyperlink(heading)
            self.sdtContent.append(p)



