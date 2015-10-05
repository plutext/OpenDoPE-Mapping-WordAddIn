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
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
//using XmlMappingTaskPane; // for Namespaces

namespace OpenDoPEModel
{
    public class Model
    {
        static Logger log = LogManager.GetLogger("Model");

        public Office.CustomXMLPart xpathsPart { get; set; }
        public Office.CustomXMLPart conditionsPart { get; set; }
        public Office.CustomXMLPart questionsPart { get; set; }
        public Office.CustomXMLPart componentsPart { get; set; }
        public Office.CustomXMLPart answersPart { get; set; }  // FabDocx only

        public IList<Office.CustomXMLPart> userParts { get; set; }

        //private Word.Document document;

        public static bool isOpenDoPEPart(Office.CustomXMLPart cp)
        {
            return cp.NamespaceURI.StartsWith("http://opendope.org/");
        }

        public static Model ModelFactory(Word.Document document)
        {
            Model model = new Model();
            model.userParts = new List<Office.CustomXMLPart>();
            //model.document = document;

            foreach (Office.CustomXMLPart cp in document.CustomXMLParts)
            {
                log.Debug("cxp: " + cp.DocumentElement + ", " + cp.NamespaceURI
                    + ", " + cp.Id);

                if (cp.NamespaceURI.Equals(Namespaces.XPATHS))
                {
                    model.xpathsPart = cp;
                }
                else if (cp.NamespaceURI.Equals(Namespaces.CONDITIONS))
                {
                    model.conditionsPart = cp;
                }
                else if (cp.NamespaceURI.Equals(Namespaces.COMPONENTS))
                {
                    model.componentsPart = cp;
                }
                else if (cp.NamespaceURI.Equals(Namespaces.QUESTIONS))
                {
                    model.questionsPart = cp;
                }
                else if (cp.NamespaceURI.Equals(Namespaces.FABDOCX_ANSWERS))
                {
                    model.answersPart = cp;
                }
                else if (!cp.BuiltIn)
                {
                    model.userParts.Add(cp);
                }
            }
            return model;
        }

        public void RemoveParts()
        {
            xpathsPart.Delete();
            xpathsPart = null;

            conditionsPart.Delete();
            conditionsPart = null;

            if (xpathsPart!=null) {
                questionsPart.Delete();
                questionsPart = null;
            }
            if (componentsPart!=null) {
                componentsPart.Delete();
                componentsPart = null;
            }

            if (answersPart != null)
            {
                answersPart.Delete();
                answersPart = null;
            }

            foreach (Office.CustomXMLPart part in userParts) {
                part.Delete();
            }
            userParts.Clear();
        }

        public Office.CustomXMLPart getUserPart(string requiredRoot) {

            if (userParts.Count == 0)
            {
                log.Error("No users part found! This shouldn't happen!");
                return null;
            }
            else
            {
                // The alternative to this would be to look
                // for a binding, and if there is one, using the
                // part that points to.

                if (string.IsNullOrWhiteSpace(requiredRoot)) {
                    // default to first
                    return userParts[0];
                } else {

                    foreach (Office.CustomXMLPart up in userParts)
                    {
                        if (up.DocumentElement.BaseName.Equals(requiredRoot))
                        {
                            return up;
                        }
                    }
                    
                    return null; // since requiredRoot is set
                }

            } 

        }


    }
}
