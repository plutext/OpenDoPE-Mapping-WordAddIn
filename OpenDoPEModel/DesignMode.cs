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

namespace OpenDoPEModel
{

    /* You can't add a content control inside one which is ShowingPlaceholderText,
     * if you are in design mode.
     * 
     * If you are not in design mode, it may work properly; the Word UI shows
     * the entire control selected.
     * 
     * In design mode, the contents of the control are not selected, and we get:
     * 
     * System.Runtime.InteropServices.COMException (0x800A11FD): This command is not available.
at Microsoft.Office.Interop.Word.ContentControls.Add(WdContentControlType Type, Object& Range)
     * 
     * ShowingPlaceholderText is a read only property, so we can't just 
     * set it to false.
     * 
     * The problem arises then if two things are true:
     * 1. the content control is ShowingPlaceholderText
     * 2. you're in design mode
     * 
     * So we can work around it by negating either of those conditions.
     * 
     * To negate the first, make a content control out of selected text.
     * ie, whenever a repeat or condition CC is added, make sure something is 
     * selected, and use that to make the content control.
     *
     * To negate the second, 
     * 
            Word.Document docx = Globals.ThisAddIn.Application.ActiveDocument;
            if (docx.FormsDesign)
            {
                docx.ToggleFormsDesign();
            }
     * 
     * The second is the easier way to go.
     * 
     * However, in practice, sometimes it doesn't seem to be enough.
     * 
     * See ribbon condition stuff for an example of negating the first.
     * 
     * There is another error which sometimes comes up "locked for editing".
     * 
        This formulation seems more susceptible to "locked for editing"
        //object rng = CurrentDocument.Application.Selection.Range;
        //cc = CurrentDocument.ContentControls.Add(Word.WdContentControlType.wdContentControlText, ref rng);

        // so prefer:
        cc = CurrentDocument.Application.Selection.ContentControls.Add(CCType, ref missing);
     */


    public class DesignMode
    {

        /// <summary>
        /// The state the user had the setting.
        /// </summary>
        private bool state;
        private Word.Document docx;

        public DesignMode(Word.Document docx)
        {
            state = docx.FormsDesign;
            this.docx = docx;
        }

        public void restoreState()
        {
            if (state)
            {
                On();
            }
            else
            {
                Off();
            }
        }

        public void Off()
        {
            if (docx.FormsDesign)
            {
                docx.ToggleFormsDesign();
            }
        }

        public void On()
        {
            if (!docx.FormsDesign)
            {
                docx.ToggleFormsDesign();
            }
        }

        // OpenDoPEModel.DesignMode.On(CurrentDocument);

    }
}
