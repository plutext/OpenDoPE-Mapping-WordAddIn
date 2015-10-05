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
//using Office = Microsoft.Office.Core;

namespace XmlMappingTaskPane
{
    public class ContentControlOpenDoPEType
    {

        public static bool isBound(Word.ContentControl cc)
        {
            if //((cc.Title != null && cc.Title.StartsWith("Data"))
                ((cc.Tag != null && cc.Tag.Contains("od:xpath"))
                || cc.XMLMapping.IsMapped
                )
            {
                return true;
            }
            return false;
        }

        public static bool isCondition(Word.ContentControl cc)
        {
            if // ((cc.Title != null && cc.Title.StartsWith("Condition"))
                 (cc.Tag != null && cc.Tag.Contains("od:condition"))
            {
                return true;
            }
            return false;

        }

        public static bool isRepeat(Word.ContentControl cc)
        {
            if //((cc.Title != null && cc.Title.StartsWith("Repeat"))
                (cc.Tag != null && cc.Tag.Contains("od:repeat"))
            {
                return true;
            }
            return false;

        }


    }
}
