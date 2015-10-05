//Copyright (c) Microsoft Corporation.  All rights reserved.
/*
 *  From http://xmlmapping.codeplex.com/license:

    Microsoft Platform and Application License

    This license governs use of the accompanying software. If you use the software, you accept this license. If you 
    do not accept the license, do not use the software.

    1. Definitions
    The terms “reproduce,” “reproduction,” “derivative works,” and “distribution” have the same meaning here as 
    under U.S. copyright law.
    A “contribution” is the original software, or any additions or changes to the software.
    A “contributor” is any person that distributes its contribution under this license.
    “Licensed patents” are a contributor’s patent claims that read directly on its contribution.

    2. Grant of Rights
    (A) Copyright Grant- Subject to the terms of this license, including the license conditions and limitations in 
    section 3, each contributor grants you a non-exclusive, worldwide, royalty-free copyright license to reproduce 
    its contribution, prepare derivative works of its contribution, and distribute its contribution or any derivative 
    works that you create.
    (B) Patent Grant- Subject to the terms of this license, including the license conditions and limitations in section 
    3, each contributor grants you a non-exclusive, worldwide, royalty-free license under its licensed patents to 
    make, have made, use, sell, offer for sale, import, and/or otherwise dispose of its contribution in the software 
    or derivative works of the contribution in the software.

    3. Conditions and Limitations
    (A) No Trademark License- This license does not grant you rights to use any contributors’ name, logo, or
    trademarks.
    (B) If you bring a patent claim against any contributor over patents that you claim are infringed by the
    software, your patent license from such contributor to the software ends automatically.
    (C) If you distribute any portion of the software, you must retain all copyright, patent, trademark, and
    attribution notices that are present in the software.
    (D) If you distribute any portion of the software in source code form, you may do so only under this license
    by including a complete copy of this license with your distribution. If you distribute any portion of the 
    software in compiled or object code form, you may only do so under a license that complies with this license.
    (E) The software is licensed “as-is.” You bear the risk of using it. The contributors give no express warranties, 
    guarantees or conditions. You may have additional consumer rights under your local laws which this license 
    cannot change. To the extent permitted under your local laws, the contributors exclude the implied warranties 
    of merchantability, fitness for a particular purpose and non-infringement.
    (F) Platform Limitation- The licenses granted in sections 2(A) & 2(B) extend only to the software or derivative
    works that you create that (1) run on a Microsoft Windows operating system product, and (2) operate with 
    Microsoft Word.
 */
using System.Globalization;
using System.Text;

namespace XmlMappingTaskPane
{
    static class ClipboardUtilities
    {

        internal static string GenerateClipboardHTML(bool needBind, string strXPath, string strPrefixMap, string strStoreId, Utilities.MappingType cctType,
            string strTitle, string strTag)
        {
            return GenerateClipboardHTML( needBind, strXPath, strPrefixMap, strStoreId, cctType, strTitle, strTag, null);
        }

        /// <summary>
        /// Generate the HTML to put on the clipboard for drag and drop.
        /// </summary>
        /// <param name="strXPath">A string specifying the XPath to the node.</param>
        /// <param name="strPrefixMap">A string specifying any necessary prefix mappings.</param>
        /// <param name="strStoreId">A string specifying the ID to the corresponding XML part in the file.</param>
        /// <param name="cctType">A MappingType specifying the type of control to insert.</param>
        /// <returns>The HTML needed for drag/drop to succeed.</returns>
        internal static string GenerateClipboardHTML(bool needBind, string strXPath, string strPrefixMap, string strStoreId, Utilities.MappingType cctType,
            string strTitle, string strTag, string strContent)
        {
            //create HTML
            StringBuilder sb = new StringBuilder();
            Encoding encoding = Encoding.UTF8;

            //build it up
            string strPrefix = string.Format(CultureInfo.InvariantCulture, cHtmlPrefix, encoding.WebName);
            string strStyles = cHtmlStyles;
            string strHtmlBody = "";
            if (needBind)
            {
                if (cctType == Utilities.MappingType.Picture)
                {
                    strHtmlBody = @"<w:Sdt PrefixMappings=""" + strPrefixMap + @""" Xpath=""" + strXPath + @""" "
                                         + GetTypeInfo(cctType) + @"StoreItemID=""" + ConvertStoreID(strStoreId)
                                         + @""" Title=""" + strTitle + @""" SdtTag=""" + strTag
                                         + @"""></w:Sdt>";
                }
                else
                {
                    strHtmlBody = @"<w:Sdt PrefixMappings=""" + strPrefixMap + @""" Xpath=""" + strXPath + @""" ShowingPlcHdr=""t"" "
                                         + GetTypeInfo(cctType) + @"StoreItemID=""" + ConvertStoreID(strStoreId)
                                         + @""" Title=""" + strTitle + @""" SdtTag=""" + strTag
                                         + @"""><p class='MsoNormal'><span lang=X-NONE><w:sdtPr></w:sdtPr></span><span class='MsoPlaceholderText'>"
                                         + GetPlaceholderText(cctType) + "</span></w:Sdt>";
                }
            }
            else if (strContent!=null)
            {
                strHtmlBody = @"<w:Sdt ShowingPlcHdr=""t"" "
                                     + GetTypeInfo(cctType) 
                                     + @""" Title=""" + strTitle + @""" SdtTag=""" + strTag
                                     + @"""><p class='MsoNormal'><span lang=X-NONE><w:sdtPr></w:sdtPr></span><span class='MsoNormal'>"
                                     + strContent + "</span></w:Sdt>";
            }
            else
            {
                strHtmlBody = @"<w:Sdt ShowingPlcHdr=""t"" "
                                     + GetTypeInfo(cctType) 
                                     + @""" Title=""" + strTitle + @""" SdtTag=""" + strTag
                                     + @"""><p class='MsoNormal'><span lang=X-NONE><w:sdtPr></w:sdtPr></span><span class='MsoPlaceholderText'>"
                                     + GetPlaceholderText(cctType) + "</span></w:Sdt>";
            }

            string strSuffix = cHtmlSuffix;

            // Get lengths of chunks
            int HeaderLength = encoding.GetByteCount(cHeader);
            HeaderLength -= 16; // extra formatting characters {0:000000}

            //determine html points
            int StartHtml = HeaderLength;
            int StartFragment = StartHtml + encoding.GetByteCount(strPrefix) + encoding.GetByteCount(strStyles);
            int EndFragment = StartFragment + encoding.GetByteCount(strHtmlBody);
            int EndHtml = EndFragment + encoding.GetByteCount(strSuffix);

            // Build the data
            sb.AppendFormat(CultureInfo.InvariantCulture, cHeader, StartHtml, EndHtml, StartFragment, EndFragment);
            sb.Append(strPrefix);
            sb.Append(strStyles);
            sb.Append(strHtmlBody);
            sb.Append(strSuffix);

            return sb.ToString();
        }

        /// <summary>
        /// Get the placeholder text for the drag/drop HTML.
        /// </summary>
        /// <param name="cctType">A MappingType specifying the type of control to insert.</param>
        /// <returns>A string specifying the corresponding placeholder text.</returns>
        private static string GetPlaceholderText(Utilities.MappingType cctType)
        {
            switch (cctType)
            {
                case Utilities.MappingType.Text:
                    return Properties.Resources.PlainTextPlaceholder;
                case Utilities.MappingType.DropDown:
                    return Properties.Resources.DropDownPlaceholder;
                case Utilities.MappingType.Picture:
                    return string.Empty;
                case Utilities.MappingType.Date:
                    return Properties.Resources.DatePlaceholder;
                default:
                    return string.Empty;
            }
        }

        /// <summary>
        /// Get the HTML declaration for the content control type being inserted.
        /// </summary>
        /// <param name="cctType">A MappingType specifying the type of control to insert.</param>
        /// <returns>A string specifying the corresponding type HTML.</returns>
        private static string GetTypeInfo(Utilities.MappingType cctType)
        {
            switch (cctType)
            {
                case Utilities.MappingType.Text:
                    return @"Text=""t"" MultiLine=""t""";
                case Utilities.MappingType.DropDown:
                    return @"DropDown=""t""";
                case Utilities.MappingType.Picture:
                    return @"DisplayAsPicture=""t""";
                case Utilities.MappingType.Date:
                    return @"Calendar=""t"" MapToDateTime=""t""";
                default:
                    return string.Empty;
            }
        }

        /// <summary>
        /// Convert the ID for the XML part into its HTML format.
        /// </summary>
        /// <param name="strOriginal">A string specifying the original ID.</param>
        /// <returns>A string specifying the HTML-equivalent ID.</returns>
        private static string ConvertStoreID(string strOriginal)
        {
            //precondition: store id is in this format: {B4F38F7D-C32D-4ADF-94A1-A1ECD73D1373}
            //postcondition: store id converted to: X_B4F38F7D-C32D-4ADF-94A1-A1ECD73D1373
            strOriginal = strOriginal.Replace("{", "X_");
            strOriginal = strOriginal.Replace("}", string.Empty);

            return strOriginal;
        }

        const string cHeader = @"Version: 1.0
StartHTML: {0:000000}
EndHTML: {1:000000}
StartFragment: {2:000000}
EndFragment: {3:000000}
";

        const string cHtmlPrefix = @"<html xmlns:o=""urn:schemas-microsoft-com:office:office""
xmlns:w=""urn:schemas-microsoft-com:office:word""
xmlns:m=""http://schemas.microsoft.com/office/2004/12/omml""
xmlns=""http://www.w3.org/TR/REC-html40"">

<head>
<meta http-equiv=Content-Type content=""text/html; charset={0}"">
<meta name=ProgId content=Word.Document>
<meta name=Generator content=""Microsoft Word 12"">
<meta name=Originator content=""Microsoft Word 12"">";

        const string cHtmlStyles = @"<style>
<!--
/* Font Definitions */
@font-face
{font-family:""Cambria Math"";
panose-1:2 4 5 3 5 4 6 3 2 4;
mso-font-charset:1;
mso-generic-font-family:roman;
mso-font-format:other;
mso-font-pitch:variable;
mso-font-signature:0 0 0 0 0 0;}
@font-face
{font-family:Calibri;
panose-1:2 15 5 2 2 2 4 3 2 4;
mso-font-charset:0;
mso-generic-font-family:swiss;
mso-font-pitch:variable;
mso-font-signature:-1610611985 1073750139 0 0 159 0;}
/* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
{mso-style-unhide:no;
mso-style-qformat:yes;
mso-style-parent:"""";
margin-top:0in;
margin-right:0in;
margin-bottom:10.0pt;
margin-left:0in;
line-height:115%;
mso-pagination:widow-orphan;
font-size:11.0pt;
font-family:""Calibri"",""sans-serif"";
mso-ascii-font-family:Calibri;
mso-ascii-theme-font:minor-latin;
mso-fareast-font-family:Calibri;
mso-fareast-theme-font:minor-latin;
mso-hansi-font-family:Calibri;
mso-hansi-theme-font:minor-latin;
mso-bidi-font-family:""Times New Roman"";
mso-bidi-theme-font:minor-bidi;}
span.MsoPlaceholderText
{mso-style-noshow:yes;
mso-style-priority:99;
mso-style-unhide:no;
color:gray;}
.MsoChpDefault
{mso-style-type:export-only;
mso-default-props:yes;
mso-ascii-font-family:Calibri;
mso-ascii-theme-font:minor-latin;
mso-fareast-font-family:Calibri;
mso-fareast-theme-font:minor-latin;
mso-hansi-font-family:Calibri;
mso-hansi-theme-font:minor-latin;
mso-bidi-font-family:""Times New Roman"";
mso-bidi-theme-font:minor-bidi;}
.MsoPapDefault
{mso-style-type:export-only;
margin-bottom:10.0pt;
line-height:115%;}
-->
</style>
</head>

<body>
<!--StartFragment-->";

        const string cHtmlSuffix = @"<!--EndFragment--></body></html>";
    }
}
