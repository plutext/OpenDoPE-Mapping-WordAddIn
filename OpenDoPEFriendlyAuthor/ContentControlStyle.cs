/*
 *  OpenDoPE authoring Word AddIn
    Copyright (C) Plutext Pty Ltd, 2012
 * 
    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using NLog;
using System.Threading;

namespace XmlMappingTaskPane
{
    /// <summary>
    /// If the adjacent runs are styled, adopt
    /// that style.
    /// 
    /// This is only necessary in the drag/drop case,
    /// since when a content control is inserted in
    /// the normal way, the desired behaviour happens
    /// automatically.
    /// </summary>
    public class ContentControlStyle
    {
        static Logger log = LogManager.GetLogger("ContentControlStyle");

        /*
         * Note the THREE places style info can appear:
         * 
             <w:sdt>
                <w:sdtPr>
                  <w:rPr>
                    <w:rStyle w:val="Heading1Char"/>
                  </w:rPr>
                  <w:id w:val="968472055"/>
                  <w:placeholder>
                    <w:docPart w:val="DefaultPlaceholder_1082065158"/>
                  </w:placeholder>
                  <w:showingPlcHdr/>
                  <w:text/>
                </w:sdtPr>
                <w:sdtEndPr>
                  <w:rPr>
                    <w:rStyle w:val="DefaultParagraphFont"/>
                    <w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>
                    <w:b w:val="0"/>
                    <w:bCs w:val="0"/>
                    <w:color w:val="FF0000"/>
                    <w:sz w:val="36"/>
                    <w:szCs w:val="22"/>
                  </w:rPr>
                </w:sdtEndPr>
         * 
         *      <w:sdtContent>
                  <w:r w:rsidRPr="00616F57">
                    <w:rPr>
                      <w:u w:val="single"/>
                    </w:rPr>
                    <w:t>ahhaaaaaaaaaa</w:t>
                  </w:r>
                </w:sdtContent>
         * 
         * A content control inserted programmatically or via Word 2010 Developer 
         * menu appears to set:
         * 
         *       <w:sdt>
                    <w:sdtPr>
                      <w:rPr>
                        <w:color w:val="FF0000"/>
                        <w:sz w:val="32"/>
                      </w:rPr>
                      <w:id w:val="1015963520"/>
                      <w:placeholder>
                        <w:docPart w:val="DefaultPlaceholder_1082065158"/>
                      </w:placeholder>
                      <w:text/>
                    </w:sdtPr>
                    <w:sdtContent>
                      <w:proofErr w:type="gramStart"/>
                      <w:r w:rsidR="00EF4BC6">
                        <w:rPr>
                          <w:color w:val="FF0000"/>
                          <w:sz w:val="32"/>
                        </w:rPr>
                        <w:t>kkk</w:t>
                      </w:r>
                      <w:proofErr w:type="gramEnd"/>
                    </w:sdtContent>
                  </w:sdt>
         * 
         * ie not w:sdtEndPr, which cc.set_DefaultTextStyle(ref style) sets
         * 
         * Tested insertion at :
         * - first character in paragraph
         * - last character in paragraph (with and without paragraph mark formatted)
         * - between 2 differently formatted runs
         * 
         * The rule seems to be that Word uses the formatting to the left,
         * except when it is at the beginning, in which case it uses the formatting
         * of the character to the right. If the character to the right is the 
         * paragraph mark, it will use that.
         * 
                */

        //public Word.ContentControl CopyAdjacentFormat(Word.ContentControl cc, bool foo) // foo arg does nothing, it is just to match delegate signature
        //{
        //    return CopyAdjacentFormat(cc);
        //}

        public Word.ContentControl CopyAdjacentFormat(Word.ContentControl cc) 
        {
            //Thread.Sleep(200);
            object collapseStart = Word.WdCollapseDirection.wdCollapseStart;
            object collapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
            object unitCharacter = Word.WdUnits.wdCharacter;

            Word.Document activeDoc = Globals.ThisAddIn.Application.ActiveDocument;

            Word.Range leftRange = cc.Range.Duplicate;
            Word.Range rightRange = cc.Range.Duplicate;

            Word.Range pRange = activeDoc.Windows[1].Selection.Paragraphs[1].Range;
            int pStartPos = pRange.Start;

            object countL = -2; // -1 is not enough
            leftRange.MoveStart(ref unitCharacter, ref countL);

            leftRange.Collapse(ref collapseStart); // it is ok for the range to be completely collapsed
            //leftRange.Text = "@";

            if (pStartPos > leftRange.Start)
            {
                // Oops, we were at the start of the paragraph,
                // so use the character to the right
                log.Info("Applying RIGHT formatting...");
                object countR = +1; // -1 is not enough
                rightRange.MoveEnd(ref unitCharacter, ref countR);
                rightRange.Collapse(ref collapseEnd);

                rightRange.Select();
            }
            else
            {
                // Usual case
                log.Info("Applying LEFT formatting...");
                leftRange.Select();
            }
            // Now apply the formatting
            activeDoc.Windows[1].Selection.CopyFormat();
            cc.Range.Select();
            activeDoc.Windows[1].Selection.PasteFormat();

            cc.Application.ScreenUpdating = true;
            cc.Application.ScreenRefresh();

            return cc;

            //WORKS
            ////cc.Range.Font = leftRange.Font;
            //cc.Range.Bold = 1;
            //cc.Range.Underline = Word.WdUnderline.wdUnderlineDotDashHeavy;
            //cc.Range.Italic = 1;

            // but DOESN'T WORK- NOT CLEAR WHY
            ////cc.Range.Font = leftRange.Font;
            //cc.Range.Bold = leftRange.Bold;
            //cc.Range.Underline = leftRange.Underline;
            //cc.Range.Italic = leftRange.Italic;

            //string te = cc.Range.Text;
            //cc.Range.FormattedText = leftRange.FormattedText;

            //Tools.PlainTextContentControl pptc = cc as Tools.PlainTextContentControl;
            
        }

    }
}
