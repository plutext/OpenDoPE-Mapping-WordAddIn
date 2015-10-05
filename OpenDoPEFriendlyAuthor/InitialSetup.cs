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
//using System.ComponentModel;
using System.Data;
//using System.Drawing;
using System.Linq;
using System.Text;

using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Word.Extensions; // for VSTOObject

using NLog;
using OpenDoPEModel;

namespace XmlMappingTaskPane
{
    class InitialSetup
    {
        static Logger log = LogManager.GetLogger("OpenDoPE_Wed");

        List<Office.CustomXMLPart> cxp = new List<Office.CustomXMLPart>();

        private void buttonOK_Click(object sender, EventArgs e)
        {
            process();
        }

        /// <summary>
        /// Create OpenDoPE parts, including optionally, question part.
        /// </summary>
        public void process()
        {

            Microsoft.Office.Interop.Word.Document document = null;
            try
            {
                document = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception ex)
            {
                Mbox.ShowSimpleMsgBoxError("No document is open/active. Create or open a docx first.");
                return;
            }

            Model model = Model.ModelFactory(document);

            // Button shouldn't be available if this exists,
            // but ..
            if (model.conditionsPart == null)
            {
                conditions conditions = new conditions();
                string conditionsXml = conditions.Serialize();
                model.conditionsPart = addCustomXmlPart(document, conditionsXml);
            }

            if (model.componentsPart == null)
            {
                components components = new components();
                string componentsXml = components.Serialize();
                model.componentsPart = addCustomXmlPart(document, componentsXml);
            }

            // Add XPath
            xpaths xpaths = new xpaths();
            // Button shouldn't be available if this exists,
            // but ..
            if (model.xpathsPart != null)
            {
                xpaths.Deserialize(model.xpathsPart.XML, out xpaths);
            }
            int idInt = 1;
            foreach (Word.ContentControl cc in Globals.ThisAddIn.Application.ActiveDocument.ContentControls)
            {
                if (cc.XMLMapping.IsMapped)
                {
                    log.Debug("Adding xpath for " + cc.ID);
                    // then we need to add an XPath
                    string xmXpath = cc.XMLMapping.XPath;

                    xpathsXpath item = new xpathsXpath();
                    // I make no effort here to check whether the xpath
                    // already exists, since the part shouldn't already exist!

                    item.id = "x" + idInt;
                    

                    xpathsXpathDataBinding db = new xpathsXpathDataBinding();
                    db.xpath = xmXpath;
                    db.storeItemID = cc.XMLMapping.CustomXMLPart.Id;
                    if (!string.IsNullOrWhiteSpace(cc.XMLMapping.PrefixMappings))
                        db.prefixMappings = cc.XMLMapping.PrefixMappings;
                    item.dataBinding = db;

                    xpaths.xpath.Add(item);

                    // Write tag
                    TagData td = new TagData(cc.Tag);
                    td.set("od:xpath", item.id);
                    cc.Tag = td.asQueryString();

                    log.Debug(".. added for " + cc.ID);
                    idInt++;
                }
            }
            string xpathsXml = xpaths.Serialize();
            if (model.xpathsPart == null)
            {
                model.xpathsPart = addCustomXmlPart(document, xpathsXml);
            }
            else
            {
                CustomXmlUtilities.replaceXmlDoc(model.xpathsPart, xpathsXml);
            }


            Microsoft.Office.Tools.Word.Document extendedDocument 
                = Globals.ThisAddIn.Application.ActiveDocument.GetVstoObject(Globals.Factory);

            //Microsoft.Office.Tools.CustomTaskPane ctp
            //    = Globals.ThisAddIn.createCTP(document, cxp, xpathsPart, conditionsPart, questionsPart, componentsPart);
            //extendedDocument.Tag = ctp;
            //// Want a 2 way association
            //WedTaskPane wedTaskPane = (WedTaskPane)ctp.Control;
            //wedTaskPane.associatedDocument = document;

            //extendedDocument.Shutdown += new EventHandler(
            //    Globals.ThisAddIn.extendedDocument_Shutdown);

            //taskPane.setupCcEvents(document);

            log.Debug("Done. Task pane now also open.");

        }


        private String getCandidatePartsNames(List<Office.CustomXMLPart> cxp)
        {
            StringBuilder sb = new StringBuilder();

            bool first = true;
            foreach (Office.CustomXMLPart cp in cxp)
            {
                Office.CustomXMLNode node = cp.SelectSingleNode("/node()");
                if (first)
                {
                    sb.Append(node.BaseName);
                    first = false;
                }
                else
                {
                    sb.Append(", " + node.BaseName);
                }
            }

            return sb.ToString();
        }

        Office.CustomXMLPart addCustomXmlPart(Word.Document document, string xml)
        {
            object missing = System.Reflection.Missing.Value;

            Office.CustomXMLPart cxp = document.CustomXMLParts.Add(xml, missing);

            log.Debug("part added");

            //bool result = cxp.LoadXML("<mynewpart><blagh/></mynewpart>");
            /* 
            * Can't do this .. causes System.Runtime.InteropServices.COMException
            * "This custom XML part has already been loaded"
            * 
            * Why?  What is the method for if it can't be used?
            * 
            * So our options are:
            * 
            * 1. replace from root node
            * 2. Delete the part, and re-add
            * 
            * Will Word remove the bindings if we do this?
            * 
            */

            //replaceXmlDoc(cxp, "<mynewpart><blagh/></mynewpart>");

            log.Debug("done");

            return cxp;
        }


    }
}
