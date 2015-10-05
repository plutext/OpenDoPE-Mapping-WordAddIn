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
using Office = Microsoft.Office.Core;  //     C:\Program Files (x86)\Microsoft Visual Studio 10.0\Visual Studio Tools for Office\PIA\Office14\Office.dll
using Word = Microsoft.Office.Interop.Word;
using NLog;
using System.Windows.Forms;

namespace OpenDoPEModel
{
    public class CustomXmlUtilities
    {
        static Logger log = LogManager.GetLogger("OpenDoPE_Wed");

        public static void replaceXmlDoc(Office.CustomXMLPart cxp, String newContent)
        {
            /* Office.CustomXMLNode node = cxp.SelectSingleNode("/");
            // gives you #document,
            // but you can't get the root node from it?

           //node = cxp.SelectSingleNode("/mypart");
             */

            Office.CustomXMLNode node = cxp.SelectSingleNode("/node()");
            //log.Debug(node.XML);
            //log.Debug(node.XPath);

            Office.CustomXMLNode parent = node.ParentNode;

            log.Debug(parent.BaseName);

            //parent.ReplaceChildNode(node, "mynewnode", "", Office.MsoCustomXMLNodeType.msoCustomXMLNodeElement, "");

            parent.ReplaceChildSubtree(newContent, node);

            //parent.RemoveChild(node); //doesn't work

            //parent.AppendChildSubtree("<mynewpart><blagh/></mynewpart>");

            node = cxp.SelectSingleNode("/");
            log.Debug(node.XML);

        }

        public static bool areOpenDoPEPartsPresent(Word.Document document)
        {
            // TODO consider removing this, or moving it,
            // refer instead Model.cs

            // Modified from OpenDoPE_Wed ThisAddIn
            Office.CustomXMLPart xpathsPart = null;
            Office.CustomXMLPart conditionsPart = null;

            foreach (Office.CustomXMLPart cp in document.CustomXMLParts)
            {
                if (cp.NamespaceURI.Equals(Namespaces.XPATHS))
                {
                    // check for duplicate parts; introduced 2013 12 08
                    if (xpathsPart != null)
                    {
                        MessageBox.Show("Duplicate parts detected. This template needs to be manually repaired.  Please contact your help desk.");                        
                    }


                    xpathsPart = cp;
                }
                else if (cp.NamespaceURI.Equals(Namespaces.CONDITIONS))
                {
                    conditionsPart = cp;
                }
            }

            return (xpathsPart != null && conditionsPart != null);

        }

        public static Office.CustomXMLPart getPartById(Word.Document document, String id) {

            foreach (Office.CustomXMLPart cp in document.CustomXMLParts)
            {
                if (cp.Id.Equals(id))
                {
                    return cp;
                }
            }
            return null;
        }

    }
}
