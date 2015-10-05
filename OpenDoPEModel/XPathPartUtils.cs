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
using Office = Microsoft.Office.Core;
//using XmlMappingTaskPane;

namespace OpenDoPEModel
{
    class XPathPartUtils
    {

        static Logger log = LogManager.GetLogger("OpenDoPE_Wed");

        /// <summary>
        /// Stuff required to create XPath element
        /// (except for ID, which this method generates)
        /// </summary>
        /// <param name="xpath"></param>
        /// <param name="storeItemID"></param>
        /// <param name="prefixMappings"></param>
        /// <param name="questionID"></param>
        public static xpathsXpath updateXPathsPart(
            Office.CustomXMLPart xpathsPart,
            string xpath, 
            string xpathId,
            string storeItemID, string prefixMappings,
                                        string questionID)
        {
            //Office.CustomXMLPart xpathsPart = ((WedTaskPane)this.Parent.Parent.Parent).xpathsPart;

            xpaths xpaths = new xpaths();
            xpaths.Deserialize(xpathsPart.XML, out xpaths);

            xpathsXpath item = new xpathsXpath();
            item.id = xpathId; //System.Guid.NewGuid().ToString(); 

            if (!string.IsNullOrWhiteSpace(questionID))
                item.questionID = questionID;

            xpathsXpathDataBinding db = new xpathsXpathDataBinding();
            db.xpath = xpath;
            db.storeItemID = storeItemID;
            if (!string.IsNullOrWhiteSpace(prefixMappings))
                db.prefixMappings = prefixMappings;
            item.dataBinding = db;

            xpaths.xpath.Add(item);

            // Save it in docx
            string result = xpaths.Serialize();
            log.Info(result);
            CustomXmlUtilities.replaceXmlDoc(xpathsPart, result);

            return item;

        }

    }
}
