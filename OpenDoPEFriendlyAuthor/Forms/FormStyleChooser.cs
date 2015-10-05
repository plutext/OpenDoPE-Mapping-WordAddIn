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
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NLog;
using Word = Microsoft.Office.Interop.Word;
using OpenDoPEModel;

namespace XmlMappingTaskPane.Forms
{



    public partial class FormStyleChooser : Form
    {

        static Logger log = LogManager.GetLogger("FormStyleChooser");

        Microsoft.Office.Interop.Word.ContentControl cc;

        Word.Style existingStyle = null;

        object styleDefaultPFont = null; //Word.Style

        /// <summary>
        /// For deselecting
        /// </summary>
        WrappedStyle selectedValue = null;

        public FormStyleChooser()
        {
            InitializeComponent();

            styleDefaultPFont = Globals.ThisAddIn.Application.ActiveDocument.Styles[Word.WdBuiltinStyle.wdStyleDefaultParagraphFont];

            cc = ContentControlMaker.getActiveContentControl(Globals.ThisAddIn.Application.ActiveDocument, Globals.ThisAddIn.Application.Selection);
            if (cc == null)
            {
                // Shouldn't happen
                log.Error("Which content control?");
                this.Close();
            }


            dynamic d = cc.get_DefaultTextStyle(); // actually returns the PARAGRAPH style!
            if (d == null)
            {
                // Nothing selected
                log.Error("No current style!");

            }
            else
            {
                //log.Debug(((Object)d).GetType().FullName);
                // System.__ComObject

                //string typeName = ComHelper.GetTypeName(d);
                //log.Debug(typeName);
                //Marshal.ReleaseComObject(selection);

                existingStyle = d as Word.Style;
                log.Debug(existingStyle.NameLocal); // eg Heading 6 ie the PARAGRAPH style!

            }

  

            foreach (Word.Style s in Globals.ThisAddIn.Application.ActiveDocument.Styles) {

                try
                {

                    if (s.Type.Equals(Word.WdStyleType.wdStyleTypeParagraph) && s.Linked)
                    {
                        //log.Debug(s.NameLocal);
                        //log.Debug(".. linked to " + linkedS.NameLocal);

                        WrappedStyle wrapped = new WrappedStyle(s);
                        listBox1.Items.Add(wrapped);

                        if (existingStyle!=null 
                            && wrapped.ToString().Equals(existingStyle.NameLocal))
                        {
                            //log.Debug(" found  " + existingStyle.NameLocal);

                            listBox1.SelectedItem = wrapped;
                        }


                    }
                }
                catch (System.Runtime.InteropServices.COMException) { }

            }



            // events click, doubleclick, mouseXXX
            //this.listBox1.SelectedIndexChanged += new System.EventHandler(listBox1_SelectedIndexChanged);
            this.listBox1.SelectedValueChanged += new System.EventHandler(listBox1_SelectedValueChanged);
        }

        /// <summary>
        /// Action when user presses cancel
        /// </summary>
        public void revert()
        {
            if (existingStyle == null)  
            {
                cc.set_DefaultTextStyle(ref styleDefaultPFont);
            } else 
            {
                object pStyle = existingStyle;
                cc.set_DefaultTextStyle(ref pStyle);
            }  
        }

        void listBox1_SelectedValueChanged(object sender, System.EventArgs e)
        {

            WrappedStyle ws = listBox1.SelectedItem as WrappedStyle;
            
            if (ws == null)
            {
                log.Warn("SelectedValueChanged, but IsNullOrEmpty.  Deselction?");
                return;
            }
            else
            {
                object pStyle;
                if (ws == selectedValue)
                {
                    // Fired on something already selected, so deselect
                    listBox1.SelectedIndex = -1; // no selection

                    cc.set_DefaultTextStyle(ref styleDefaultPFont); // can't set to null?

                    selectedValue = null;
                    return;
                }

                Word.Style selectedStyle = ws.getStyle();
                pStyle = ws.getStyle();
                // Displayed the paragraph style; use the linked character style 
                // (but setting the p style seems to work)
                log.Info("selected " + selectedStyle.NameLocal);
                //object newStyle = (Word.Style)selectedStyle.get_LinkStyle();
                cc.set_DefaultTextStyle(ref pStyle);  // I wonder whether we can use the p style?

                selectedValue = ws; //for next time

                // Now, select cc, and apply p style, so user see visible result
                // (only need to do that if you set c style (as opposed to p)
                // But it only applies the style to the first paragraph, so:
                cc.Range.set_Style(ref pStyle);

            }

        }

        //void listBox1_SelectedIndexChanged(object sender, System.EventArgs e)
        //{
        //    if (listBox1.SelectedItems.Count == 0)
        //    {
        //        log.Warn("SelectedIndexChanged, but nothing selected.  Deselction?");
        //        return;
        //    }

        //    int selectedItemIndex = listBox1.SelectedIndex;
        //    string selectedItemText = listBox1.SelectedItem.ToString();
        //    // eg [Balloon Text Char, Balloon Text]
        //    log.Info("selectedItemText: " + selectedItemText);
        //}


    }

    /* 
     * Dictionary<Word.Style, string> selectableStyles = new Dictionary<Word.Style, string>();
     * 
     * BindingSource bs = new BindingSource(selectableStyles, null);
       listBox1.DataSource = bs; 
       listBox1.DisplayMember = "Value";
       listBox1.ValueMember = "Key";
     * 
     * approach does display list entries,
     * but can't seem to programmatically set the selected style
     * whether key is style or stylename
     * 
     * Not sure whether I tried SelectedItem (how would you do that for a dictionary entry?); 
     * I certainly tried selectedValue
     * 
     * so better to avoid BindingSource!!!
     */

    public class WrappedStyle
    {
        Word.Style s;
        public WrappedStyle(Word.Style s)
        {
            this.s = s;
        }

        public override string ToString()
        {
            return s.NameLocal;
        }

        public Word.Style getStyle()
        {
            return s;
        }
    }
}
