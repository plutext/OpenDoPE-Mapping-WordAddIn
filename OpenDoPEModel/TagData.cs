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
using System.Collections.Specialized;

namespace OpenDoPEModel
{
    public class TagData
    {
        // Represents data using query string format,  
        // but doesn't bother to URL encode/decode.
        // Anything not in key=value format
        // is stored as a key with a null value

        Dictionary<string, string> map; // or NameValueCollection?

        public TagData(string queryString)
        {

            map = new Dictionary<string, string>();

            if (string.IsNullOrEmpty(queryString)) return;

            string[] querySegments = queryString.Split('&');
            foreach (string segment in querySegments)
            {
                string[] parts = segment.Split('=');
                if (parts.Length > 1)
                {
                    string key = parts[0];
                    string val = parts[1];

                    map.Add(key, val);
                }
                else
                {
                    // Not a key value pair; represent as key
                    map.Add(segment, null);
                }
            }
        }

        public string get(string name)
        {
            try
            {
                return map[name];
            }
            catch (KeyNotFoundException)
            {                
                return null;
            }
        }

        public string getXPathID()
        {
            try
            {
                return map["od:xpath"];
            }
            catch (KeyNotFoundException)
            {
                return null;
            }
        }

        public string getRepeatID()
        {
            try
            {
                return map["od:repeat"];
            }
            catch (KeyNotFoundException)
            {
                return null;
            }
        }

        public string getConditionID()
        {
            try
            {
                return map["od:condition"];
            }
            catch (KeyNotFoundException)
            {
                return null;
            }
        }
        public void set(string name, string val)
        {
            remove(name);
            map.Add(name, val);
        }

        public void remove(string name)
        {
            try
            {
                map.Remove(name);
            }
            catch (KeyNotFoundException)
            { }
        }


        public String asQueryString()
        {

            StringBuilder sb = new StringBuilder();

            int pos = 0;
            foreach (string key in map.Keys)
            {
                if (pos > 0)
                {
                    sb.Append("&");
                }
                if (map[key] == null)
                {
                    sb.Append(key);
                }
                else
                {
                    sb.Append(key + "=" + map[key]);
                }
                pos++;
            }

            return sb.ToString();

        }		

    }
}
