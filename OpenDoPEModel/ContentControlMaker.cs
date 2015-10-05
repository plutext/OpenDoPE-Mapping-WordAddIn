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
using NLog;

using Word = Microsoft.Office.Interop.Word;
//using Office = Microsoft.Office.Core;
//using XmlMappingTaskPane;

namespace OpenDoPEModel
{
    public class ContentControlMaker
    {
        static Logger log = LogManager.GetLogger("OpenDoPE_Wed");

        /// <summary>
        /// Create a new CC if not already in one.
        /// </summary>
        /// <param name="controlType"></param>
        /// <param name="docx"></param>
        /// <param name="selection"></param>
        /// <returns></returns>
        public static Word.ContentControl MakeOrReuse(bool tryToReuse, Word.WdContentControlType controlType, 
            Word.Document docx, Word.Selection selection)
        {

            // step 1: create content control, if necessary
            // are we in a content control?
            Word.ContentControl cc = null;
            if (tryToReuse)
            {
                cc = ContentControlMaker.getActiveContentControl(docx, selection);
            }
            object missing = System.Type.Missing;
            if (cc == null)
            {
                // Add one
                cc = docx.ContentControls.Add(controlType, ref missing);
                cc.Title = "[New]";
            }
            else
            {
                // we're in a content control already.
                // Have they selected all of it?  
                // If so, is this add or re-map?
                if (controlType.Equals(Word.WdContentControlType.wdContentControlRichText))
                {
                    cc.XMLMapping.Delete();
                }
                cc.Type = controlType;
            }

            return cc;

        }

        /// <summary>
        /// Return the content control containing the cursor, or
        /// null if we are outside a content control or the selection
        /// spans the content control boundary
        /// </summary>
        /// <param name="docx"></param>
        /// <param name="selection"></param>
        /// <returns></returns>
        public static Word.ContentControl getActiveContentControl(Word.Document docx, Word.Selection selection)
        {
            //Word.Selection selection = Globals.ThisAddIn.Application.Selection;

            // Word.ContentControls ccs = selection.ContentControls;
            // only has a value if the selection contains one or both *ends* of a ContentControl

            // so how do you expand a selection to include the entire content control?
            // or can we ask whether a content control is active,
            // ie the selection is in it?

            // or iterate through the content controls in the active doc,
            // asking whether their range contains my range? YES
            // Hmmm... what about nesting of cc?

            Word.ContentControl result = null;

            foreach (Word.ContentControl ctrl in docx.ContentControls)
            {
                int highestCount = -1;
                if (selection.InRange(ctrl.Range))
                {
                    //diagnostics("DEBUG - Got control");
                    log.Debug("In control: " + ctrl.ID);
                    int parents = countParent(ctrl);
                    if (parents > highestCount)
                    {
                        log.Debug(".. highest ");
                        result = ctrl;
                        highestCount = parents;
                    }
                }
                // else user's selection is totally outside the content control, or crosses its boundary
            }
            //if (stateDocx.InControl == true)

            return result;

        }

        private static int countParent(Word.ContentControl ctrl)
        {
            int counter = 0;

            while (ctrl.ParentContentControl != null)
            {
                counter++;
                ctrl = ctrl.ParentContentControl;
            }
            log.Debug("= " + counter);
            return counter;
        }


    }

}
