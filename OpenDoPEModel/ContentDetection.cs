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

namespace OpenDoPEModel
{
    public class ContentDetection
    {
        static Logger log = LogManager.GetLogger("RibbonMapping");

        public static bool IsBase64Encoded(String str)
        {
            log.Debug("Testing " + str);

            // See http://www.codeproject.com/Questions/177808/How-to-determine-if-a-string-is-Base64-decoded-or
            try
            {
                // If no exception is caught, then it is possibly a base64 encoded string
                byte[] data = Convert.FromBase64String(str);

                // But "Figs" doesn't throw an exception.
                // Smallest gif is 35 byte; png 67; jpeg 125.
                if (data.Length < 35)
                {
                    return false;
                }
                else
                {
                    //log.Debug("Length " + data.Length);
                }

                // The part that checks if the string was properly padded to the
                // correct length was borrowed from d@anish's solution
                return (str.Replace(" ", "").Length % 4 == 0);
            }
            catch (Exception e)
            {
                // If exception is caught, then it is not a base64 encoded string
                log.Debug("Caught  " + e.Message + "; so not base64 encoded");
                return false;
            }
        }

        /// <summary>
        /// FIXME: the Add-In has already unescaped content, so there is no &lt; ?
        /// </summary>
        /// <param name="content"></param>
        /// <returns></returns>
        public static bool IsXHTMLContent(String content)
        {
            log.Debug("inspecting " + content);
            return content.Contains("</div") || content.Contains("</span") || content.Contains("</p") || content.Contains("</li") || content.Contains("</td");
        }

        public static bool IsFlatOPCContent(String content)
        {
            //log.Debug("inspecting " + content);
            return content.Contains("pkg:package");
        }

    }
}
