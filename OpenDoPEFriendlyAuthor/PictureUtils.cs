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
using OpenDoPEModel;

namespace XmlMappingTaskPane
{
    public class PictureUtils
    {

        public static void pastePictureIntoCC(Word.ContentControl cc, byte[] picBytes)
        {

            //log.Info(".. here is where we insert image ");

            System.IO.MemoryStream ms = new System.IO.MemoryStream(picBytes);
            System.Drawing.Image img = System.Drawing.Image.FromStream(ms);

            System.Windows.Forms.Clipboard.SetImage(img);  // bitmap

            cc.Range.Paste();  // nb, this line fails if you try to do it in ContentControlAfterAdd event, so do all this here...

            ///*
            // * Options:
            // * 
            // * 1. save/load from disk
            // * 2. write to clipboard (without disk?), then paste
            // * 3. InsertXML (Sentient.WordML PictWriter)
            //   4.  Microsoft.Office.Tools.Word.PictureContentControl
            // */ 
        }

        public static void setPictureHandler(TagData td)
        {
            td.set("od:Handler", "picture");
        }


    }
}
