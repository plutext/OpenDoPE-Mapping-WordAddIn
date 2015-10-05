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
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Schema;
using Microsoft.Win32;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using OpenDoPEModel;
using NLog;
using XmlMappingTaskPane.Controls;

namespace XmlMappingTaskPane
{
    class OpenDopeCreateMappedControl
    {

        static Logger log = LogManager.GetLogger("OpenDopeCreateMappedControl");

        /// <summary>
        /// Create a content control mapped to the selected XML node.
        /// </summary>
        /// <param name="CCType">A WdContentControlType value specifying the type of control to create.</param>
        public void CreateMappedControl(Word.WdContentControlType CCType, ControlTreeView.OpenDopeType odType,
            ControlTreeView controlTreeView,
            ControlMain controlMain,
            Word.Document CurrentDocument,
            Office.CustomXMLPart CurrentPart,
            //XmlDocument OwnerDocument,
            bool _PictureContentControlsReplace
            )
        {
            OpenDoPEModel.DesignMode designMode = new OpenDoPEModel.DesignMode(CurrentDocument);
            designMode.Off();

            try
            {
                object missing = Type.Missing;
                TreeNode tn = controlTreeView.treeView.SelectedNode;
                if (((XmlNode)tn.Tag).NodeType == XmlNodeType.Text)
                {
                    tn = tn.Parent;
                }

                //get an nsmgr
                NameTable nt = new NameTable();

                //generate the xpath and the ns manager
                XmlNamespaceManager xmlnsMgr = new XmlNamespaceManager(nt);
                string strXPath = Utilities.XpathFromXn(CurrentPart.NamespaceManager, (XmlNode)tn.Tag, true, xmlnsMgr);
                log.Info("Right click for XPath: " + strXPath);

                string prefixMappings = Utilities.GetPrefixMappings(xmlnsMgr);

                // Insert bind | condition | repeat
                // depending on which mode button is pressed.
                TagData td = new TagData("");
                if ((controlMain.modeControlEnabled == false && odType == ControlTreeView.OpenDopeType.Unspecified) // ie always mode bind
                    || (controlMain.modeControlEnabled == true && controlMain.controlMode1.isModeBind())
                    || odType == ControlTreeView.OpenDopeType.Bind)
                {
                    log.Debug("In bind mode");
                    String val = ((XmlNode)tn.Tag).InnerText;

                    //bool isXHTML = HasXHTMLContent(tn);
                    bool isPicture = false;
                    bool isXHTML = false;
                    bool isFlatOPC = ContentDetection.IsFlatOPCContent(val);


                    Word.ContentControl cc = null;

                    if (isFlatOPC)
                    {
                        // <?mso-application progid="Word.Document"?>
                        // <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">

                        log.Debug(".. contains block content ");

                        cc = CurrentDocument.Application.Selection.ContentControls.Add(
                            Word.WdContentControlType.wdContentControlRichText, ref missing);

                        // Ensure block level
                        Inline2Block i2b = new Inline2Block();
                        cc = i2b.convertToBlockLevel(cc, false, true);

                        if (cc == null)
                        {
                            MessageBox.Show("Problems inserting block level WordML at this location.");
                            return;
                        }

                    }
                    else
                    {
                        isXHTML = ContentDetection.IsXHTMLContent(val);
                    }

                    if (isXHTML)
                    {
                        cc = CurrentDocument.Application.Selection.ContentControls.Add(
                            Word.WdContentControlType.wdContentControlRichText, ref missing);
                        if (Inline2Block.containsBlockLevelContent(val))
                        {
                            Inline2Block i2b = new Inline2Block();
                            cc = i2b.convertToBlockLevel(cc, true, true);

                            if (cc == null)
                            {
                                MessageBox.Show("Problems inserting block level XHTML at this location.");
                                designMode.restoreState();
                                return;
                            }
                        }

                    }
                    else if (ContentDetection.IsBase64Encoded(val))
                    {
                        isPicture = true;

                        if (_PictureContentControlsReplace)
                        {
                            // Use a rich text control instead
                            cc = CurrentDocument.ContentControls.Add(
                                Word.WdContentControlType.wdContentControlRichText, ref missing);

                            PictureUtils.pastePictureIntoCC(cc, Convert.FromBase64String(val));
                        }
                        else
                        {
                            // Force picture content control
                            log.Debug("Detected picture");
                            cc = CurrentDocument.Application.Selection.ContentControls.Add(Word.WdContentControlType.wdContentControlPicture, ref missing);
                        }
                    }
                    else if (!isFlatOPC)
                    {
                        log.Debug("Not picture or XHTML; " + CCType.ToString());

                        // This formulation seems more susceptible to "locked for editing"
                        //object rng = CurrentDocument.Application.Selection.Range;
                        //cc = CurrentDocument.ContentControls.Add(Word.WdContentControlType.wdContentControlText, ref rng);

                        // so prefer:
                        cc = CurrentDocument.Application.Selection.ContentControls.Add(CCType, ref missing);

                    }

                    XPathsPartEntry xppe = new XPathsPartEntry(controlMain.model);
                    xppe.setup(null, CurrentPart.Id, strXPath, prefixMappings, true);
                    xppe.save();

                    td.set("od:xpath", xppe.xpathId);

                    if (isFlatOPC)
                    {
                        td.set("od:progid", "Word.Document");
                        cc.Title = "Word: " + xppe.xpathId;
                        //cc.Range.Text = val; // don't escape it
                        cc.Range.InsertXML(val, ref missing);


                    }
                    else if (isXHTML)
                    {
                        td.set("od:ContentType", "application/xhtml+xml");
                        cc.Title = "XHTML: " + xppe.xpathId;
                        cc.Range.Text = val;
                    }
                    else if (isPicture)
                    {
                        PictureUtils.setPictureHandler(td);
                        cc.Title = "Image: " + xppe.xpathId;

                    }
                    else
                    {
                        cc.Title = "Data value: " + xppe.xpathId;
                    }
                    cc.Tag = td.asQueryString();

                    if (cc.Type == Word.WdContentControlType.wdContentControlText)
                    {
                        cc.MultiLine = true;
                    }
                    if (cc.Type != Word.WdContentControlType.wdContentControlRichText)
                    {
                        cc.XMLMapping.SetMappingByNode(Utilities.MxnFromTn(tn, CurrentPart, true));
                    }

                    designMode.restoreState();
                }
                else if ((controlMain.modeControlEnabled == true && controlMain.controlMode1.isModeCondition())
                    || odType == ControlTreeView.OpenDopeType.Condition)
                {
                    log.Debug("In condition mode");

                    // User can make a condition whatever type they like,
                    // but if they make it text, change it to RichText.
                    if (CCType == Word.WdContentControlType.wdContentControlText)
                    {
                        CCType = Word.WdContentControlType.wdContentControlRichText;
                    }
                    Word.ContentControl cc = CurrentDocument.Application.Selection.ContentControls.Add(CCType, ref missing);
                    ConditionsPartEntry cpe = new ConditionsPartEntry(controlMain.model);
                    cpe.setup(CurrentPart.Id, strXPath, prefixMappings, true);
                    cpe.save();

                    cc.Title = "Conditional: " + cpe.conditionId;
                    // Write tag
                    td.set("od:condition", cpe.conditionId);
                    cc.Tag = td.asQueryString();

                    // We want to be in Design Mode, so user can see their gesture take effect
                    designMode.On();

                    Ribbon.editXPath(cc);
                }
                else if ((controlMain.modeControlEnabled == true && controlMain.controlMode1.isModeRepeat())
                    || odType == ControlTreeView.OpenDopeType.Repeat)
                {
                    log.Debug("In repeat mode");


                    // User can make a repeat whatever type they like
                    // (though does it ever make sense for it to be other than RichText?),
                    // but if they make it text, change it to RichText.
                    if (CCType == Word.WdContentControlType.wdContentControlText)
                    {
                        CCType = Word.WdContentControlType.wdContentControlRichText;
                    }
                    Word.ContentControl cc = CurrentDocument.Application.Selection.ContentControls.Add(CCType, ref missing);

                    // Need to drop eg [1] (if any), so BetterForm-based interactive processing works
                    if (strXPath.EndsWith("]"))
                    {
                        strXPath = strXPath.Substring(0, strXPath.LastIndexOf("["));
                        log.Debug("Having dropped '[]': " + strXPath);
                    }

                    XPathsPartEntry xppe = new XPathsPartEntry(controlMain.model);
                    xppe.setup("rpt", CurrentPart.Id, strXPath, prefixMappings, false);
                    xppe.save();
                    cc.Title = "Data value: " + xppe.xpathId;
                    // Write tag
                    td.set("od:repeat", xppe.xpathId);
                    cc.Tag = td.asQueryString();

                    // We want to be in Design Mode, so user can see their gesture take effect
                    designMode.On();
                }
            }
            catch (COMException cex)
            {
                controlTreeView.ShowErrorMessage(cex.Message);
                designMode.restoreState();
            }
        }

    }
}
