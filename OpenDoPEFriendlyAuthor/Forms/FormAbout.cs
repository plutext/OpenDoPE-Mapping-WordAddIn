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

namespace XmlMappingTaskPane.Forms
{
    public partial class FormAbout : Form
    {
        public FormAbout()
        {
            InitializeComponent();

            webBrowser1.Navigate("about:blank");
            HtmlDocument doc = this.webBrowser1.Document;
            doc.Write(String.Empty);

            this.webBrowser1.DocumentText = "<html><body><p><b>OpenDoPE authoring Word Add-In, version 1.2, Jan 2021</b></p>"
                + "<p>This basic authoring tool is designed to support merging of XML data at run time.</p>"
                + "<p>Compared to our FabDocx authoring tool, this authoring tool:</p>"
                +"<ul>"
                + "<li>supports your choice of XML data format, and shows you that XML</li>"
                + "<li>automates simple conditions</li>"
                + "<li>does not create questions for interactive use (the main FabDocx use case)</li>"
                + "</ul>"
                + "<p>For discussion, try <a href='http://www.docx4java.org/forums/data-binding-java-f16/'>this forum</a></p>"
                + "<p>(c) Copyright 2013-2021 Plutext Pty Ltd, portions copyright Microsoft Corporation</p></body></html>";

        }
    }
}
