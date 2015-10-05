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

namespace XmlMappingTaskPane
{
    public class ContentControlUtilities
    {

        static Logger log = LogManager.GetLogger("ContentControlUtilities");


        public static List<Word.ContentControl> getShallowestSelectedContentControls(Word.Document docx, Word.Selection selection)
        {
            // Word.ContentControls ccs = selection.ContentControls;
            // only has a value if the selection contains one or both *ends* of a ContentControl

            // so how do you expand a selection to include the entire content control?
            // or can we ask whether a content control is active,
            // ie the selection is in it?

            // or iterate through the content controls in the active doc,
            // asking whether their range contains my range? YES
            // Hmmm... what about nesting of cc?

            List<Word.ContentControl> results = new List<Word.ContentControl>();

            int lowestCount = 99;
            foreach (Word.ContentControl ctrl in docx.ContentControls)
            {
                if (ctrl.Range.InRange(selection.Range))
                {
                    //diagnostics("DEBUG - Got control");
                    log.Debug("In control: " + ctrl.ID);
                    int parents = countParent(ctrl);
                    if (parents < lowestCount)
                    {
                        log.Debug(".. lowest ");
                        lowestCount = parents;
                    }
                }
            }

            if (lowestCount == 99)
            {
                // No content control was entirely selected.
                return results; // empty list
            }

            foreach (Word.ContentControl ctrl in docx.ContentControls)
            {
                if (ctrl.Range.InRange(selection.Range))
                {
                    //diagnostics("DEBUG - Got control");
                    log.Debug("In control: " + ctrl.ID);
                    int parents = countParent(ctrl);
                    if (parents == lowestCount)
                    {
                        results.Add(ctrl);
                    }
                }
            }
            return results;
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
