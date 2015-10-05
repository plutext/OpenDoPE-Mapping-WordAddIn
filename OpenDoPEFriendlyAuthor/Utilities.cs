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
using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Schema;
using Microsoft.Office.Tools;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using NLog;

namespace XmlMappingTaskPane
{
    static class Utilities
    {

        static Logger log = LogManager.GetLogger("Utilities");

        internal enum MappingType { Text, Date, DropDown, Picture, RichText };

        /// <summary>
        /// Convert an XmlNodeType into the corresponding CustomXMLNodeType.
        /// </summary>
        /// <param name="xmlNodeType">The input XmlNodeType.</param>
        /// <returns>The corresponding CustomXMLNodeType.</returns>
        internal static Office.MsoCustomXMLNodeType CxntFromXnt(XmlNodeType xmlNodeType)
        {
            switch (xmlNodeType)
            {
                case XmlNodeType.Element:
                    return Microsoft.Office.Core.MsoCustomXMLNodeType.msoCustomXMLNodeElement;
                case XmlNodeType.Attribute:
                    return Microsoft.Office.Core.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute;
                case XmlNodeType.CDATA:
                    return Microsoft.Office.Core.MsoCustomXMLNodeType.msoCustomXMLNodeCData;
                case XmlNodeType.ProcessingInstruction:
                    return Microsoft.Office.Core.MsoCustomXMLNodeType.msoCustomXMLNodeProcessingInstruction;
                case XmlNodeType.Text:
                    return Microsoft.Office.Core.MsoCustomXMLNodeType.msoCustomXMLNodeText;
                case XmlNodeType.Comment:
                    return Microsoft.Office.Core.MsoCustomXMLNodeType.msoCustomXMLNodeComment;
            }

            throw new ArgumentException("invalid node type");
        }

        /// <summary>
        /// Get the prefix mappings needed to evaluate the XPath to a specific custom XML node.
        /// </summary>
        /// <param name="cxn">The corresponding CustomXMLNode.</param>
        /// <returns>A string specifying the prefix mapping list.</returns>
        internal static string GetPrefixMappingsMxn(Office.CustomXMLNode cxn)
        {
            string s = "";
            foreach (Office.CustomXMLPrefixMapping cxpm in cxn.OwnerPart.NamespaceManager)
            {
                //get the string
                if (!String.IsNullOrEmpty(cxpm.Prefix) && cxpm.Prefix != "xml" && cxpm.Prefix != "xmlns")
                    s += "xmlns:" + cxpm.Prefix + "='" + cxpm.NamespaceURI + "' ";
            }
            return s;
        }

        /// <summary>
        /// Get a CustomXMLNode from a TreeNode.
        /// </summary>
        /// <param name="tn">The TreeNode to convert.</param>
        /// <param name="cxp">The CustomXMLPart containing the corresponding CustomXMLNode.</param>
        /// <param name="fRemoveTextNode">True to get the parent XML node, False otherwise.</param>
        /// <returns>The corresponding CustomXMLNode.</returns>
        internal static Office.CustomXMLNode MxnFromTn(TreeNode tn, Office.CustomXMLPart cxp, bool fRemoveTextNode)
        {
            if (tn == null || tn.Text == "/")
                throw new ArgumentNullException("tn");

            //if we hit a null node, bail
            if (((XmlNode)tn.Tag) == null)
                throw new ArgumentNullException("tn");

            //get an nsmgr
            NameTable nt = new NameTable();
            XmlNamespaceManager xmlnsMgr = new XmlNamespaceManager(nt);

            //check if we're editing the text node of an attribute, since then we'll want to get the XPath of the attribute
            string xpath;
            if (((XmlNode)tn.Tag).NodeType == XmlNodeType.Text && ((XmlNode)tn.Parent.Tag).NodeType == XmlNodeType.Attribute)
            {
                xpath = XpathFromTn(tn.Parent, fRemoveTextNode, cxp, xmlnsMgr);
            }
            else if (((XmlNode)tn.Tag).NodeType == XmlNodeType.Text)
            {
                xpath = XpathFromTn(tn, fRemoveTextNode, cxp, xmlnsMgr);
            }
            else
            {
                xpath = XpathFromTn(tn, false, cxp, xmlnsMgr);
            }

            Debug.Assert(!String.IsNullOrEmpty(xpath), "ASSERT: empty xpath", "xpathFromTn gave us back nothing!");

            Office.CustomXMLNode selectedNode = cxp.SelectSingleNode(xpath);

            Debug.Assert(selectedNode != null, "ASSERT: null mxn from xpath", "This XPath: " + xpath + " gave us back no node!");
            return selectedNode;
        }

        /// <summary>
        /// Get an XPath from a TreeNode.
        /// </summary>
        /// <param name="tn">The TreeNode to convert.</param>
        /// <param name="removeTextNode">True to get the parent XML node's XPath, False otherwise.</param>
        /// <param name="cxp">The CustomXMLPart containing the corresponding CustomXMLNode.</param>
        /// <param name="xnsmgr">The XmlNamespaceManager for the local XML part.</param>
        /// <returns>The XPath to the XML node.</returns>
        internal static string XpathFromTn(TreeNode tn, bool removeTextNode, Office.CustomXMLPart cxp, XmlNamespaceManager xnsmgr)
        {
            return Utilities.XpathFromXn(cxp.NamespaceManager, tn.Tag as XmlNode, removeTextNode, xnsmgr);
        }

        /// <summary>
        /// Get an XmlNode in the local XML tree from the corresponding CustomXMLNode.
        /// </summary>
        /// <param name="xdoc">The XmlDocument for the local XML tree.</param>
        /// <param name="mxn">The CustomXMLNode to convert.</param>
        /// <param name="mxnNew">A new CustomXMLNode that's not in the local tree, and therefore should be ignored when trying to find the XmlNode.</param>
        /// <returns>The corresponding XmlNode.</returns>
        internal static XmlNode XnFromMxn(XmlDocument xdoc, Office.CustomXMLNode mxn, Office.CustomXMLNode mxnNew)
        {
            NameTable nt = new NameTable();
            XmlNamespaceManager xmlnsMgr = new XmlNamespaceManager(nt);
            string strXPath = string.Empty;
            XmlNode xn = null;

            strXPath = XpathFromMxn(mxn, xmlnsMgr, mxnNew);
            Debug.Assert(!string.IsNullOrEmpty(strXPath));

            if (xdoc != null)
            {
                xn = xdoc.SelectSingleNode(strXPath, xmlnsMgr);
            }

            Debug.Assert(xn != null);

            return xn;
        }

        /// <summary>
        /// Create an XmlNode for a CustomXMLNode.
        /// </summary>
        /// <param name="mxnNewNode">The CustomXMLNode to convert.</param>
        /// <param name="xdoc">The XmlDocument for the corresponding XmlNode.</param>
        /// <returns>The newly created XmlNode.</returns>
        internal static XmlNode XnBuildFromMxn(Office.CustomXMLNode mxnNewNode, XmlDocument xdoc)
        {
            XmlNode xn = null;
            switch (mxnNewNode.NodeType)
            {
                case Office.MsoCustomXMLNodeType.msoCustomXMLNodeElement:
                    //create a new temp XML document
                    XmlDocument xdocElement = new XmlDocument();
                    xdocElement.LoadXml(mxnNewNode.XML);

                    //import the element back into the main DOM
                    xn = xdoc.ImportNode(xdocElement.DocumentElement, true);

                    //clean up
                    xdocElement = null;
                    break;
                case Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute:
                    char[] cSplitter = { '=' };
                    string[] strAttribute = mxnNewNode.XML.Split(cSplitter);
                    xn = xdoc.CreateAttribute(strAttribute[0], mxnNewNode.NamespaceURI);
                    xn.InnerText = mxnNewNode.NodeValue;
                    break;
                case Office.MsoCustomXMLNodeType.msoCustomXMLNodeText:
                    xn = xdoc.CreateTextNode(mxnNewNode.Text);
                    break;
                case Office.MsoCustomXMLNodeType.msoCustomXMLNodeProcessingInstruction:
                    xn = xdoc.CreateProcessingInstruction(mxnNewNode.BaseName, mxnNewNode.Text);
                    break;
                case Office.MsoCustomXMLNodeType.msoCustomXMLNodeCData:
                    xn = xdoc.CreateCDataSection(mxnNewNode.Text);
                    break;
                case Office.MsoCustomXMLNodeType.msoCustomXMLNodeComment:
                    xn = xdoc.CreateComment(mxnNewNode.Text);
                    break;
            }
            return xn;
        }

        /// <summary>
        /// Get an XPath for a CustomXMLNode.
        /// </summary>
        /// <param name="xn">The CustomXMLNode to convert.</param>
        /// <param name="xnsmgr">The XmlNamespaceManager for the local XML tree.</param>
        /// <param name="mxnNew">A new CustomXMLNode that's not in the local tree, and therefore should be ignored when trying to find the XPath.</param>
        /// <returns>The corresponding XPath.</returns>
        private static string XpathFromMxn(Office.CustomXMLNode xn, XmlNamespaceManager xnsmgr, Office.CustomXMLNode mxnNew)
        {
            try
            {
                Debug.Assert(xn != null, "cannot create an xpath from a null node");

                string strThis = null;
                string strThisName = null;
                string ns = "";

                bool fCheckParent = true;

                Office.CustomXMLNode xnParent = xn.ParentNode;                
                Office.CustomXMLPrefixMappings mnsmgr = xn.OwnerPart.NamespaceManager;

                switch (xn.NodeType)
                {
                    case Office.MsoCustomXMLNodeType.msoCustomXMLNodeElement:
                        strThisName = xn.BaseName;

                        if (!string.IsNullOrEmpty(xn.NamespaceURI))
                        {
                            ns = mnsmgr.LookupPrefix(xn.NamespaceURI) + ":";
                            xnsmgr.AddNamespace(mnsmgr.LookupPrefix(xn.NamespaceURI), xn.NamespaceURI);
                        }
                        strThis = "/" + ns + strThisName + XpathPosFromMxn(xn, mxnNew);
                        break;

                    case Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute:
                        strThisName = xn.BaseName;
                        if (!string.IsNullOrEmpty(xn.NamespaceURI))
                        {
                            ns = mnsmgr.LookupPrefix(xn.NamespaceURI) + ":";
                            xnsmgr.AddNamespace(mnsmgr.LookupPrefix(xn.NamespaceURI), xn.NamespaceURI);
                        }
                        strThis = "/@" + ns + strThisName;
                        xnParent = xn.SelectSingleNode("..");

                        break;

                    case Office.MsoCustomXMLNodeType.msoCustomXMLNodeProcessingInstruction:
                        {
                            strThis = "/processing-instruction(";

                            strThis = strThis + ")" + XpathPosFromMxn(xn, mxnNew);
                            break;
                        }

                    case Office.MsoCustomXMLNodeType.msoCustomXMLNodeText:
                        strThis = "/text()" + XpathPosFromMxn(xn, mxnNew);
                        break;

                    case Office.MsoCustomXMLNodeType.msoCustomXMLNodeComment:
                        strThis = "/comment()" + XpathPosFromMxn(xn, mxnNew);
                        break;

                    case Office.MsoCustomXMLNodeType.msoCustomXMLNodeDocument:
                        strThis = "";
                        fCheckParent = false;
                        break;

                    case Office.MsoCustomXMLNodeType.msoCustomXMLNodeCData:
                        strThis = "/text()" + XpathPosFromMxn(xn, mxnNew);
                        break;
                }

                return strThis.Insert(0, fCheckParent ? XpathFromMxn(xnParent, xnsmgr, mxnNew) : "");
            }
            catch (COMException ex)
            {
                Debug.Fail(ex.Source, ex.Message);
            }

            return string.Empty;
        }

        /// <summary>
        /// Find the position of a node at its level in the XML part.
        /// </summary>
        /// <param name="xn">The XmlNode to convert.</param>
        /// <param name="mxnNew">A new CustomXMLNode that's not in the local tree, and therefore should be ignored when trying to find the position.</param>
        /// <returns>The position of the node, in XPath format (i.e. [1]).</returns>
        private static string XpathPosFromMxn(Office.CustomXMLNode xn, Office.CustomXMLNode mxnNew)
        {
            try
            {
                Debug.Assert(xn != null, "cannot create an xpath from a null node");

                long lSib = 0;
                long lTotal = 0;
                Office.CustomXMLNode xnPrevSib = null;
                string name = xn.BaseName;
                string nsUri = xn.NamespaceURI;
                Office.MsoCustomXMLNodeType nt = xn.NodeType;

                if (xn == null)
                {
                    Debug.Fail("null node");
                    throw new ArgumentNullException("xn");
                }

                if (xn.ParentNode != null)
                    lTotal = xn.ParentNode.ChildNodes.Count;

                //need to get the mxnNew down all the way to here and NOT increment the count 
                //*if* the xnPrevSib == mxnNew, since then the new node is
                //*not* in my DOM, and I should not use it in the XPath(!)
                while ((xnPrevSib = xn.PreviousSibling) != null)
                {
                    if (mxnNew != null)
                    {
                        if (xnPrevSib.NodeType == nt &&
                            xnPrevSib.BaseName.Equals(name) &&
                            xnPrevSib.NamespaceURI.Equals(nsUri) &&
                            xnPrevSib.XPath != mxnNew.XPath)
                            lSib++;
                    }
                    else
                    {
                        if (xnPrevSib.NodeType == nt &&
                            xnPrevSib.BaseName.Equals(name) &&
                            xnPrevSib.NamespaceURI.Equals(nsUri))
                            lSib++;
                    }

                    xn = xnPrevSib;
                }

                return lTotal > 0 ? "[" + (lSib + 1).ToString(CultureInfo.InvariantCulture) + "]" : "";
            }
            catch (ArgumentNullException ex)
            {
                Debug.Fail(ex.Source, ex.Message);
            }

            return string.Empty;
        }

        /// <summary>
        /// Get an XPath for an XML node in the local XML tree.
        /// </summary>
        /// <param name="mnsmgr">A CustomXMLPrefixMappings containing the namespace mappings in the corresponding CustomXMLPart.</param>
        /// <param name="xn">The XmlNode to convert.</param>
        /// <param name="removeTextNode">True to get the XPath for the parent XML node, False otherwise.</param>
        /// <param name="xnsmgr">The XmlNamespaceManager for the local XML tree.</param>
        /// <returns>The corresponding XPath.</returns>
        internal static string XpathFromXn(Office.CustomXMLPrefixMappings mnsmgr, XmlNode xn, bool removeTextNode, XmlNamespaceManager xnsmgr)
        {
            string strThis = null;
            string strThisName = null;
            string ns = "";
            XmlNode xnParent = xn.ParentNode;
            bool fCheckParent = true;

            switch (xn.NodeType)
            {
                case XmlNodeType.Element:
                    strThisName = xn.LocalName;

                    if (!String.IsNullOrEmpty(xn.NamespaceURI) && xnsmgr != null && mnsmgr != null && !string.IsNullOrEmpty(mnsmgr.LookupPrefix(xn.NamespaceURI)))
                    {
                        xnsmgr.AddNamespace(mnsmgr.LookupPrefix(xn.NamespaceURI), xn.NamespaceURI);
                        ns = mnsmgr.LookupPrefix(xn.NamespaceURI) + ":";
                    }
                    strThis = "/" + ns + strThisName + XpathPosFromXn(xn);
                    break;

                case XmlNodeType.Attribute:
                    strThisName = xn.LocalName;
                    if (!String.IsNullOrEmpty(xn.NamespaceURI) && xnsmgr != null && mnsmgr != null)
                    {
                        xnsmgr.AddNamespace(mnsmgr.LookupPrefix(xn.NamespaceURI), xn.NamespaceURI);
                        ns = mnsmgr.LookupPrefix(xn.NamespaceURI) + ":";
                    }
                    strThis = "/@" + ns + strThisName;
                    xnParent = xn.SelectSingleNode("..", null);

                    break;

                case XmlNodeType.ProcessingInstruction:
                    {
                        strThis = "/processing-instruction(";

                        strThis = strThis + ")" + XpathPosFromXn(xn);
                        break;
                    }

                case XmlNodeType.Text:
                    //if the parent is an attribute (for this text node), we need to get it's parent's xpath instead
                    if (xn.ParentNode.NodeType != XmlNodeType.Attribute && !removeTextNode)
                    {
                        strThis = "/text()" + XpathPosFromXn(xn);
                    }
                    else
                    {
                        strThis = "";
                    }
                    break;

                case XmlNodeType.Comment:
                    strThis = "/comment()" + XpathPosFromXn(xn);
                    break;

                case XmlNodeType.Document:
                    strThis = "";
                    fCheckParent = false;
                    break;

                case XmlNodeType.EntityReference:
                case XmlNodeType.CDATA:
                    strThis = "/text()" + XpathPosFromXn(xn);
                    break;

                case XmlNodeType.Whitespace:
                case XmlNodeType.SignificantWhitespace:
                    break;
            }

            return strThis.Insert(0, fCheckParent ? XpathFromXn(mnsmgr, xnParent, false, xnsmgr) : "");
        }

        /// <summary>
        /// Find the position of a node at its level in the local XML tree.
        /// </summary>
        /// <param name="xn">The XmlNode to convert.</param>
        /// <returns>The position of the node, in XPath format (i.e. [1]).</returns>
        private static string XpathPosFromXn(XmlNode xn)
        {
            Debug.Assert(xn != null, "cannot create an xpath from a null node");

            long lSib = 0;
            long lTotal = 0;
            XmlNode xnPrevSib = null;
            string name = xn.LocalName;
            string nsUri = xn.NamespaceURI;
            XmlNodeType nt = xn.NodeType;

            if (xn.ParentNode != null)
                lTotal = xn.ParentNode.ChildNodes.Count;

            while ((xnPrevSib = xn.PreviousSibling) != null)
            {
                if (xnPrevSib.NodeType == nt &&
                    xnPrevSib.LocalName.Equals(name) &&
                    xnPrevSib.NamespaceURI.Equals(nsUri))
                    lSib++;

                xn = xnPrevSib;
            }

            return lTotal > 0 ? "[" + (lSib + 1).ToString(CultureInfo.InvariantCulture) + "]" : "";
        }

        /// <summary>
        /// Get the type of content control that should be created by default for this XML node.
        /// </summary>
        /// <param name="xmlNode">An XmlNode to convert.</param>
        /// <returns>A MapingType enumeration value with the type of content control that should be created.</returns>
        internal static MappingType CheckNodeType(XmlNode xmlNode)
        {
            if (xmlNode.SchemaInfo.SchemaElement != null && xmlNode.SchemaInfo.SchemaElement.ElementSchemaType != null)
            {
                //is it xsd:dateTime or xsd:base64Binary
                switch (xmlNode.SchemaInfo.SchemaElement.ElementSchemaType.TypeCode)
                {
                    case XmlTypeCode.Date:
                    case XmlTypeCode.DateTime:
                        return MappingType.Date;
                    case XmlTypeCode.Base64Binary:
                        return MappingType.Picture;
                    default:
                        break;
                }

                //is there an enumeration?
                if (xmlNode.SchemaInfo.SchemaElement.ElementSchemaType is XmlSchemaSimpleType)
                {
                    if (((XmlSchemaSimpleType)xmlNode.SchemaInfo.SchemaElement.ElementSchemaType).Content is XmlSchemaSimpleTypeRestriction)
                    {
                        XmlSchemaSimpleTypeRestriction xsstr = (XmlSchemaSimpleTypeRestriction)((XmlSchemaSimpleType)xmlNode.SchemaInfo.SchemaElement.ElementSchemaType).Content;
                        foreach (XmlSchemaFacet xsf in xsstr.Facets)
                        {
                            if (xsf is XmlSchemaEnumerationFacet)
                                return MappingType.DropDown;
                        }
                    }
                }
            }
            else if (xmlNode.SchemaInfo.SchemaAttribute != null && xmlNode.SchemaInfo.SchemaAttribute.AttributeSchemaType != null)
            {
                //is it xsd:dateTime or xsd:base64Binary
                switch (xmlNode.SchemaInfo.SchemaAttribute.AttributeSchemaType.TypeCode)
                {
                    case XmlTypeCode.Date:
                    case XmlTypeCode.DateTime:
                        return MappingType.Date;
                    case XmlTypeCode.Base64Binary:
                        return MappingType.Picture;
                    default:
                        break;
                }

                //is there an enumeration?
                if (xmlNode.SchemaInfo.SchemaAttribute.AttributeSchemaType is XmlSchemaSimpleType)
                {
                    if (((XmlSchemaSimpleType)xmlNode.SchemaInfo.SchemaAttribute.AttributeSchemaType).Content is XmlSchemaSimpleTypeRestriction)
                    {
                        XmlSchemaSimpleTypeRestriction xsstr = (XmlSchemaSimpleTypeRestriction)((XmlSchemaSimpleType)xmlNode.SchemaInfo.SchemaAttribute.AttributeSchemaType).Content;
                        foreach (XmlSchemaFacet xsf in xsstr.Facets)
                        {
                            if (xsf is XmlSchemaEnumerationFacet)
                                return MappingType.DropDown;
                        }
                    }
                }
            }
            return MappingType.Text;
        }

        /// <summary>
        /// Get the prefix mappings used by the local XML tree.
        /// </summary>
        /// <param name="xmlnsMgr">The XmlNamespaceManager for the local XML tree.</param>
        /// <returns>A string containing all of the prefix mappings, in standard XML format (i.e. xmlns:ns0='...').</returns>
        internal static string GetPrefixMappings(XmlNamespaceManager xmlnsMgr)
        {
            if (xmlnsMgr != null)
            {
                string s = "";
                foreach (string strNS in xmlnsMgr)
                {
                    //get the string
                    if (!String.IsNullOrEmpty(strNS) && strNS != "xml" && strNS != "xmlns")
                        s += "xmlns:" + strNS + "='" + xmlnsMgr.LookupNamespace(strNS) + "' ";
                }
                return s;
            }
            return string.Empty;
        }

        /// <summary>
        /// Find the task pane for the current Word window.
        /// </summary>
        /// <returns>The CustomTaskPane for the corresponding task pane instance, or null if none exists.</returns>
        internal static CustomTaskPane FindTaskPaneForCurrentWindow()
        {
            CustomTaskPane ctpPaneForThisWindow = null;

            //look for the right one
            if (Globals.ThisAddIn.Application.ShowWindowsInTaskbar)
            {
                // ie a document frame window for each open document
                // (there is a ctp for each document frame window)
                Globals.ThisAddIn.TaskPaneList.TryGetValue(
                    Globals.ThisAddIn.Application.ActiveWindow, out ctpPaneForThisWindow);
            }
            else
            {
                Debug.Assert(Globals.ThisAddIn.CustomTaskPanes.Count <= 1, 
                    "why are there stray CTPs?");
                if (Globals.ThisAddIn.CustomTaskPanes.Count > 0)
                {
                    ctpPaneForThisWindow = Globals.ThisAddIn.CustomTaskPanes[0];
                }
            }

            bool gotCTP = ctpPaneForThisWindow != null;
            log.Debug("CTP found? " + gotCTP);

            if (gotCTP)
            {
                try
                {
                    int foo = ctpPaneForThisWindow.Height;
                }
                catch (System.ObjectDisposedException)
                {
                    // presumably the user has pressed DeleteAll,
                    // and is now trying to add XML again.
                    return null;
                }
            }

            return ctpPaneForThisWindow;
        }

        /// <summary>
        /// Get the placeholder text that should be created by default for this content control type.
        /// </summary>
        /// <param name="ccType">A WdContentControlType enumeration value specifying the content control type.</param>
        /// <returns>The corresponding placeholder text.</returns>
        internal static string GetPlaceholderText(Microsoft.Office.Interop.Word.WdContentControlType ccType)
        {
            switch (ccType)
            {
                case Word.WdContentControlType.wdContentControlText:
                    return Properties.Resources.PlainTextPlaceholder;
                case Word.WdContentControlType.wdContentControlDropdownList:
                    return Properties.Resources.DropDownPlaceholder;
                case Word.WdContentControlType.wdContentControlDate:
                    return Properties.Resources.DatePlaceholder;
                case Word.WdContentControlType.wdContentControlComboBox:
                    return Properties.Resources.DropDownPlaceholder;
                default:
                    Debug.Fail("unknown content control type");
                    throw new ArgumentOutOfRangeException("ccType");
            }
        }
    }
}
