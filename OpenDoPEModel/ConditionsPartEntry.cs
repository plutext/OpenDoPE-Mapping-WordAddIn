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
    public class ConditionsPartEntry
    {
        static Logger log = LogManager.GetLogger("ConditionsPartEntry");

        public ConditionsPartEntry(Model model)
        {
            this.model = model;

            conditions tmpconditions = new conditions();
            conditions.Deserialize(model.conditionsPart.XML, out tmpconditions);
            conditions = tmpconditions;
        }

        public conditions conditions { get; set; }
        Model model;

        public string conditionId { get; set; }

        /// <summary>
        /// Add xpath to xpaths part, and a condition to the conditions part.  
        /// </summary>
        /// <param name="model"></param>
        /// <param name="cxpId"></param>
        /// <param name="strXPath"></param>
        /// <param name="prefixMappings"></param>
        public condition setup(string cxpId, string strXPath, string prefixMappings,
            bool setupQuestionNow)
        {
            //////////////////////////
            // First, add the new XPath

            // Drop any trailing "/" from a Condition XPath
            if (strXPath.EndsWith("/"))
            {
                strXPath = strXPath.Substring(0, strXPath.Length - 1);
                log.Debug("truncated to " + strXPath);
            }

            XPathsPartEntry xppe = new XPathsPartEntry(model);
            xpathsXpath xpath = xppe.setup("cond", cxpId, strXPath, prefixMappings, setupQuestionNow);
            xppe.save();

            return setup(xpath);
        }
        /// <summary>
        /// Add the condition to the conditions part.  
        /// </summary>
        /// <param name="model"></param>
        /// <param name="cxpId"></param>
        /// <param name="strXPath"></param>
        /// <param name="prefixMappings"></param>
        public condition setup(xpathsXpath xpath)
        {

            condition result = null; 

            //////////////////////////
            // Second, create and add the condition

            // If the Condition is already defined in our Condition part, don't do it again.
            // Also need this for ID generation.


            Dictionary<string, string> conditionsById = new Dictionary<string, string>();
            foreach (condition xx in conditions.condition)
            {
                conditionsById.Add(xx.id, "");
                if (xx.Item is xpathref)
                {
                    xpathref ex = (xpathref)xx.Item;

                    if (ex.id.Equals(xpath.id))
                    {
                        result = xx;
                        log.Info("This Condition is already setup, with ID: " + xx.id);
                        break;
                    }
                }
            }



            if (result == null) // not already defined
            {
                // Add the new condition
                result = new condition();
                //result.id = IdGenerator.generateIdForXPath(conditionsById, null, null, xpath.dataBinding.xpath);
                result.id = IdHelper.GenerateShortID(5);

                xpathref xpathref = new xpathref();
                xpathref.id = xpath.id;

                result.Item = xpathref;

                conditions.condition.Add(result);

                // Save the conditions in docx
                string ser = conditions.Serialize();
                log.Info(ser);
                CustomXmlUtilities.replaceXmlDoc(model.conditionsPart, ser);
            }

            // Set this
            conditionId = result.id;

            log.Debug("Condition written!");

            return result;
        }

        /// <summary>
        /// Allocate the condition an ID, and add it
        /// </summary>
        /// <param name="condition"></param>
        public void add(condition condition, string suggestedId)
        {
            //Dictionary<string, string> conditionsById = new Dictionary<string, string>();
            //foreach (condition xx in conditions.condition)
            //{
            //    conditionsById.Add(xx.id, "");
            //}

            //condition.id = IdGenerator.generateIdForXPath(conditionsById, null, null, suggestedId);
            condition.id = IdHelper.GenerateShortID(5);

            conditions.condition.Add(condition);

            save();

            log.Debug("Condition written!");
        }

        public void save()
        {
            // Save the conditions in docx
            string ser = conditions.Serialize();
            log.Info(ser);
            CustomXmlUtilities.replaceXmlDoc(model.conditionsPart, ser);
        }


        public condition getConditionByID(String id)
        {

            return conditions.getConditionByID(id);

            //if (conditions.condition != null)
            //{
            //    for (int i = 0; i < conditions.condition.Length; i++)
            //    {
            //        condition xx = conditions.condition[i];
            //        if (xx.id.Equals(id))
            //        {
            //            return xx;
            //        }
            //    }
            //}
            //return null;
        }


    }
}
