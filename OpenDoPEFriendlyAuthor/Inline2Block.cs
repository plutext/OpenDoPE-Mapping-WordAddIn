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
using System.Windows.Forms;
using NLog;

using Word = Microsoft.Office.Interop.Word;
//using Office = Microsoft.Office.Core;

namespace XmlMappingTaskPane
{
    class Inline2Block
    {

        static Logger log = LogManager.GetLogger("ContentControlUtilities");

        Word.Application wdApp;


        static string[] blockElements = { "html", "body", "div", "table", "img", "p", "h1", "h2", "h3", "h4", "h5", "h6" };

        public static bool containsBlockLevelContent(string xhtml)
        {
            for (int i = 0; i < blockElements.Length; i++)
            {
                string el = "<" + blockElements[i];
                if (xhtml.Contains(el)) return true;
            }
            return false;
        }

        public bool isBlockLevel(Word.ContentControl currentCC)
        {
            /*
             * from start of P
                p  S
                cc S+1

                exc pmark:
                p  Z
                cc Z-2

                inc p mark:
                p   Z
                cc  Z
             * 
             *  Index of pmark = index of start of next P.
                */

            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Range ccRange = currentCC.Range;
            ccRange.Select();
            Word.Range pRange = document.Windows[1].Selection.Paragraphs[1].Range;

            log.Info("p {0}, {1}", pRange.Start, pRange.End);
            log.Info("cc {0}, {1}", ccRange.Start, ccRange.End);

            return (ccRange.End >= pRange.End);

        }

        private bool ccStartsAtStartOfParagraph(Word.ContentControl currentCC)
        {
            /*
             * from start of P
                p  S
                cc S+1

                exc pmark:
                p  Z
                cc Z-2

                inc p mark:
                p   Z
                cc  Z
             * 
             *  Index of pmark = index of start of next P.
                */

            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Range ccRange = currentCC.Range;
            ccRange.Select();
            Word.Range pRange = document.Windows[1].Selection.Paragraphs[1].Range;

            log.Info("p {0}, {1}", pRange.Start, pRange.End);
            log.Info("cc {0}, {1}", ccRange.Start, ccRange.End);

            return (pRange.Start == (ccRange.Start+1));

        }

        /// <summary>
        /// Convert a rich text control to block level, then inject WordML into it.
        /// </summary>
        /// <param name="control"></param>
        public Word.ContentControl blockLevelFlatOPC(Word.ContentControl currentCC, String xml)
        {
            Word.ContentControl cc = convertToBlockLevel(currentCC, false, false);

            object missing = System.Type.Missing;
            cc.Range.InsertXML(xml, ref missing);

            cc.Application.ScreenUpdating = true;  // screen seems to be updating already (Word 2010 x64). bugger.
            cc.Application.ScreenRefresh();

            return cc;
        }

        /// <summary>
        /// An inline rich text control can't contain carriage returns.
        /// </summary>
        /// <param name="control"></param>
        public Word.ContentControl convertToBlockLevel(Word.ContentControl currentCC, bool keepContents, bool updateScreen)
        {
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            OpenDoPEModel.DesignMode designMode = new OpenDoPEModel.DesignMode(document);

            // Only do it if the content control is rich text
            if (!currentCC.Type.Equals(Word.WdContentControlType.wdContentControlRichText))
            {
                log.Warn("convert to block level only operates on rich text controls, not " + currentCC.Type);
                return null;
            }

            string majorVersionString = Globals.ThisAddIn.Application.Version.Split(new char[] { '.' })[0];
            int majorVersion = Convert.ToInt32(majorVersionString);


            // Only do it if the content control is not already block level
            //
            if (isBlockLevel(currentCC)) return currentCC;


            // Can only do this if the content control is not
            // nested within some other inline content control
            if (currentCC.ParentContentControl != null
                && !isBlockLevel(currentCC.ParentContentControl))
            {
                MessageBox.Show("This content control contains block level content, but can't be converted automatically.  Please correct this yourself, by deleting it, and re-creating at block level.");
                return null;
            }


            if (majorVersion >= 14)
            {
                getWordApp().UndoRecord.StartCustomRecord("Promote content control to block-level");
            }

            bool ccIsAtPStart = ccStartsAtStartOfParagraph(currentCC);

            object collapseStart = Word.WdCollapseDirection.wdCollapseStart;
            //object collapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
            object unitCharacter = Word.WdUnits.wdCharacter;


            // Get a range start of cc.
            Word.Range ccRange = currentCC.Range;
            ccRange.Collapse(ref collapseStart); 

            // Delete the cc, but preserve Tag, Title
            string tagVal = currentCC.Tag;
            string titleVal = currentCC.Title;
            string contents = currentCC.Range.Text;

            currentCC.Delete(true);

            ccRange.Select();
            if (ccIsAtPStart)
            {
                document.Windows[1].Selection.TypeParagraph();
            }
            else
            {
                // Insert 2 new paragraphs
                document.Windows[1].Selection.TypeParagraph();
                document.Windows[1].Selection.TypeParagraph();
            }

            // Create a cc around the first new p
            object start = ccRange.Start+1;
            object end = ccRange.Start+1; 
            object newRange = document.Range(ref start, ref end);
            log.Info("target {0}, {1}", ((Word.Range)newRange).Start, ((Word.Range)newRange).End);

            designMode.Off();
            Word.ContentControl newCC = document.ContentControls.Add(Word.WdContentControlType.wdContentControlRichText, ref newRange);
            designMode.restoreState();

            newCC.Tag = tagVal;
            newCC.Title = titleVal;
            if (keepContents) // want to do this for XHTML
            {
                newCC.Range.Text = contents;
            }

            if (updateScreen)
            {
                newCC.Application.ScreenUpdating = true;
                newCC.Application.ScreenRefresh();
            }
            return newCC;

            // Approach:
            //  .. Get the paragraph
            // .. Make a copy
            // .. in the copy, delete up to our position
            // .. in the original, delete after our position

            //Word.Range splittingPoint = sel.Range;

            //object para1DeleteStartPoint = splittingPoint.Start;

            //Word.Range paraOrig = Globals.ThisAddIn.Application.ActiveDocument.Range(ref para1DeleteStartPoint, ref para1DeleteStartPoint);
            //paraOrig.MoveStart(ref unitParagraph, ref back1);

            //int lengthStartSegment = paraOrig.End - paraOrig.Start;

            //object startPoint = paraOrig.Start;

            //paraOrig.MoveEnd(ref unitParagraph, ref forward1);

            //object endPoint = paraOrig.End;
            //object endPointPlusOne = paraOrig.End + 1;

            //// copy it
            //Word.Range insertPoint = Globals.ThisAddIn.Application.ActiveDocument.Range(ref endPoint, ref endPoint);
            //paraOrig.Copy();
            //insertPoint.Paste();

            //// In the copy, delete the first half
            //// (do this operation first, to preserve our original position calculations)
            //object para2DeleteEndpoint = (int)endPoint + lengthStartSegment;
            //Word.Range para2Deletion = Globals.ThisAddIn.Application.ActiveDocument.Range(ref endPoint, ref para2DeleteEndpoint);
            //para2Deletion.Delete();

            //// In the original, delete the second half
            //Word.Range para1Deletion = Globals.ThisAddIn.Application.ActiveDocument.Range(ref para1DeleteStartPoint, ref endPoint);
            //para1Deletion.Delete();

            if (majorVersion >= 14)
            {
                getWordApp().UndoRecord.EndCustomRecord();
            }

        }

        Word.Application getWordApp()
        {
            if (wdApp == null)
            {
                //wdApp = new Word.Application();
                wdApp = Globals.ThisAddIn.Application;
            }
            return wdApp;
        }

    }
}
