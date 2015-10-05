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
using Office = Microsoft.Office.Core;
using NLog;

namespace OpenDoPEModel
{

    public class XPathsPartEntry
    {
        static Logger log = LogManager.GetLogger("XPathsPartEntry");

        public XPathsPartEntry(Model model)
        {
            this.model = model;

            xpaths = new xpaths();
            xpaths.Deserialize(model.xpathsPart.XML, out xpaths);
        }

        public xpaths xpaths;
        public xpaths getXPaths()
        {
            return xpaths;
        }

        Model model;


        public string xpathId { get; set; }

        /// <summary>
        /// Add the XPath to the XPaths part.  As a side effect, set fields tag
        /// and xpathid
        /// </summary>
        /// <param name="model"></param>
        /// <param name="cxpId">storeItemID for data binding</param>
        /// <param name="strXPath"></param>
        /// <param name="prefixMappings"></param>
        public xpathsXpath setup(string idTypeSuffix, string cxpId, string strXPath, 
            string prefixMappings,
            bool setupQuestionNow)
        {
            xpathsXpath result = null;

            // If the XPath is already defined in our XPaths part, don't do it again.
            // Also need this for ID generation.
            Dictionary<string, string> xpathsById = new Dictionary<string, string>();
            foreach (xpathsXpath xx in xpaths.xpath)
            {
                try
                {
                    xpathsById.Add(xx.id, "");
                }
                catch (Exception e)
                {
                    log.Error(xx.id + " exists already!"); // How did this happen??
                    throw e;
                }
                if (xx.dataBinding.xpath.Equals(strXPath)
                    && xx.dataBinding.storeItemID.Equals(cxpId))
                {
                    result = xx;
                    xpathId = xx.id;
                    log.Info("This XPath is already setup, with ID: " + xpathId);
                    break;
                }
            }

            if (result==null) // not already defined
            {
                //xpathId = IdGenerator.generateIdForXPath(xpathsById, "", idTypeSuffix, strXPath);
                xpathId = IdHelper.GenerateShortID(5);

                // Question
                string questionID = null;
                if (setupQuestionNow 
                    && model.questionsPart != null)
                {
                    throw new NotImplementedException();
                    //FormQuestion formQuestion = new FormQuestion(model.questionsPart,
                    //    strXPath, xpathId);
                    //formQuestion.ShowDialog();
                    //// TODO - handle cancel

                    //formQuestion.updateQuestionsPart(formQuestion.getQuestion());

                    //questionID = formQuestion.textBoxQID.Text;
                    //formQuestion.Dispose();
                }

                // Also add to XPaths
                result = createXpath(
                    strXPath,
                    xpathId,
                    cxpId, prefixMappings, questionID);
            }

            return result;

        }

        public xpathsXpath createXpath(
            string xpath,
            string xpathId,
            string storeItemID, string prefixMappings,
                                        string questionID)
        {
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

            return item;

        }


        public xpathsXpath getXPathByID(String id)
        {
            foreach (xpathsXpath xx in xpaths.xpath)
            {
                if (xx.id.Equals(id))
                {
                    return xx;
                }

            }
            return null;
        }

        public xpathsXpath getXPathByQuestionID(String qid)
        {
            return xpaths.getXPathByQuestionID(qid);
        }

        public void save()
        {
            // Save it in docx
            string result = xpaths.Serialize();
            log.Info(result);
            CustomXmlUtilities.replaceXmlDoc(model.xpathsPart, result);
        }

    }
}
