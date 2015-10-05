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

namespace OpenDoPEModel
{
    /// <summary>
    /// NOT USED ANYMORE,  
    /// since:
    /// 1. need to keep the ID short, given max Tag length
    /// 2. want it to be unique across documents, for easy re-use
    /// So see instead IdHelper.
    /// </summary>
    public class IdGenerator
    {

        public static string dropNonAlpha(string incoming) {

            StringBuilder sb = new StringBuilder();
            char[] chars = incoming.ToCharArray();

            for (int i = 0; i < chars.Length; i++)
            {
                char c = chars[i];
                if (char.IsLetter(c))
                {
                    // Not letters and numbers .. consider count(foo)>1
                    sb.Append(c);
                }
                else
                {
                    sb.Append(" ");
                }
            }

            String result = sb.ToString();
            return result.Trim();

    }

        /// <summary>
        /// Generate an ID which is unique (not in the dictionary supplied).
        /// </summary>
        /// <param name="xpathsPart"></param>
        /// <param name="partPrefix">if the user has several custom xml parts, pass some prefix so the names can distinguish </param>
        /// <param name="strXPath"></param>
        /// <returns></returns>
        public static string generateIdForXPath(Dictionary<string, string> xpathsById,
            string partPrefix, string typeSuffix, string strXPath)
        {
            string name = dropNonAlpha(strXPath);
            // get last segment of XPath
            int pos = name.LastIndexOf(" ");
            if (pos > 0)
            {
                name = name.Substring(pos + 1);
            }
            if (partPrefix != null && !partPrefix.Equals(""))
                name = partPrefix + "_" + name;
            if (typeSuffix != null && !typeSuffix.Equals(""))
                name = name + "_" + typeSuffix;

            if (!idExists(name, xpathsById))
            {
                return name;
            }
            // Now just increment a number until it is unique
            int i = 1;
            do
            {
                i++;
            } while (idExists(name + i, xpathsById));
            return name + i;
        }

        private static bool idExists(string id, Dictionary<string, string> xpathsById)
        {
            try
            {
                string foo = xpathsById[id];
                return true;
            }
            catch (KeyNotFoundException)
            {
                return false;
            }
        }
    }
}
