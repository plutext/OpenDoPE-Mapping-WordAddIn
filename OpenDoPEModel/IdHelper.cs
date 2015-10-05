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
using System.Text.RegularExpressions;

using NLog;

namespace OpenDoPEModel
{
    /// <summary>
    /// Used for question and answer
    /// (conditions and xpaths use ID
    ///  derived from Q/A id, shorted, given tag length restrictions!).
    /// </summary>
    public class IdHelper
    {
        static Logger log = LogManager.GetLogger("IdHelper");

        static string bannedWords = @"\b(please|enter|type|what|is|the|a|do|you|your|was|and|or)\b";

        public static String SuggestID(string suggestion, List<string> reserved)
        {
            // Want it to be an NCName (non-colonized), in case,
            // some time in the future, we need some namespace concept.

            // The practical restrictions of NCName are that it cannot contain 
            // several symbol characters like :, @, $, %, &, /, +, ,, ;, 
            // whitespace characters or different parenthesis. Furthermore 
            // an NCName cannot begin with a number, dot or minus 
            // character although they can appear later in an NCName.
            // No spaces or colons. Allows "_" and "-".

            // First, strip common words (eg, please, enter, type, what, is, the)
            suggestion = Regex.Replace(suggestion, bannedWords, "", 
                            RegexOptions.IgnoreCase);

            // Now truncate it, to 20 to 30 chars
            suggestion = suggestion.Trim();
            if (suggestion.Length > 20)
            {
                // try not to truncate middle of word
                if (suggestion.Length > 30)
                {
                    suggestion = suggestion.Substring(0, 30) + " ";
                }
                else
                {
                    suggestion = suggestion + " ";
                }
                int pos = suggestion.IndexOf(" ", 20);

                log.Debug(pos);
                if (pos > 0 && pos < 30)
                {
                    suggestion = suggestion.Substring(0, pos);
                    suggestion = suggestion.TrimEnd();
                    log.Debug(suggestion);
                }
                else
                {
                    suggestion = suggestion.Substring(0, 20);
                    suggestion = suggestion.TrimEnd();
                    log.Debug(suggestion);
                }
            }
            // .. replace spaces with "_"
            suggestion = suggestion.Replace(" ", "_");

            // .. strip out nasty characters
            // http://stackoverflow.com/questions/10938749/regex-to-strip-characters-except-given-ones
            //First replace removes all the characters you don't want.
            suggestion = Regex.Replace(suggestion, "[^_a-zA-Z0-9-]", "");
            //Second replace removes any characters from the start that aren't allowed there.
            suggestion = Regex.Replace(suggestion, "^[^a-zA-Z]+", "");

            // Finally, ensure it is unique against the passed in list
            if (reserved.Contains(suggestion))
            {
                int i = 1;
                do
                {
                    i++;
                } while (reserved.Contains(suggestion + i));
                suggestion = suggestion + i;
            }

            return suggestion;
        }

        static string charPool = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890";
        static int poolLength = charPool.Length;
        private static Random _random = new Random();

        static string alphaPool = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";

        /// <summary>
        /// 1. need to keep the ID short, given max Tag length
        /// 2. want it to be unique across documents, for easy re-use
        /// 62 possible characters x 5 length, means ~64^(5/2)=64^2.5=~32K
        /// would be required for 50% chance of collision (the birthday 
        /// paradox).  Seems acceptable.
        /// </summary>
        /// <param name="length"></param>
        /// <returns></returns>
        public static string GenerateShortID(int length)
        {
            StringBuilder rs = new StringBuilder();

            // First character: not a number
            rs.Append(alphaPool[(int)(_random.Next(alphaPool.Length - 1))]);

            for (int i = 1; i < length; i++)
            {
                rs.Append(charPool[(int)(_random.Next(poolLength - 1))]);
            }

            return rs.ToString();
        }
    }
}
