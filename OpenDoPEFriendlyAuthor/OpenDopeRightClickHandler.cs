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
    class OpenDopeRightClickHandler
    {

        static Logger log = LogManager.GetLogger("OpenDopeRightClickHandler");

        /// <summary>
        /// used when they right click then select "map to"
        /// </summary>
        /// <param name="odType"></param>
        public void mapToSelectedControl(ControlTreeView.OpenDopeType odType,
            ControlTreeView controlTreeView,
            ControlMain controlMain,
            Word.Document CurrentDocument,
            Office.CustomXMLPart CurrentPart,
            //XmlDocument OwnerDocument,
            bool _PictureContentControlsReplace
            )

        {
            object missing = System.Type.Missing;

            DesignMode designMode = new OpenDoPEModel.DesignMode(CurrentDocument);
            // In this method, we're usually not creating a control,
            // so we don't need to turn off

            try
            {
                //create a binding     
                Word.ContentControl cc = null;
                if (CurrentDocument.Application.Selection.ContentControls.Count == 1)
                {
                    log.Debug("CurrentDocument.Application.Selection.ContentControls.Count == 1");
                    object objOne = 1;
                    cc = CurrentDocument.Application.Selection.ContentControls.get_Item(ref objOne);
                    log.Info("Mapped content control to tree view node " + controlTreeView.treeView.SelectedNode.Name);
                }
                else if (CurrentDocument.Application.Selection.ParentContentControl != null)
                {
                    log.Debug("ParentContentControl != null");
                    cc = CurrentDocument.Application.Selection.ParentContentControl;
                }
                if (cc != null)
                {

                    TreeNode tn = controlTreeView.treeView.SelectedNode;

                    //get an nsmgr
                    NameTable nt = new NameTable();

                    //generate the xpath and the ns manager
                    XmlNamespaceManager xmlnsMgr = new XmlNamespaceManager(nt);
                    string strXPath = Utilities.XpathFromXn(CurrentPart.NamespaceManager, (XmlNode)tn.Tag, true, xmlnsMgr);
                    log.Info("Right clicked with XPath: " + strXPath);

                    string prefixMappings = Utilities.GetPrefixMappings(xmlnsMgr);

                    // Insert bind | condition | repeat
                    // depending on which mode button is pressed.
                    TagData td = new TagData("");
                    if ((controlMain.modeControlEnabled == false && odType == ControlTreeView.OpenDopeType.Unspecified) // ie always mode bind
                        || (controlMain.modeControlEnabled == true && controlMain.controlMode1.isModeBind())
                        || odType == ControlTreeView.OpenDopeType.Bind)
                    {
                        log.Debug("In bind mode");

                        XPathsPartEntry xppe = new XPathsPartEntry(controlMain.model);
                        xppe.setup(null, CurrentPart.Id, strXPath, prefixMappings, true);
                        xppe.save();

                        td.set("od:xpath", xppe.xpathId);

                        String val = ((XmlNode)tn.Tag).InnerText;
                        bool isXHTML = false;
                        bool isFlatOPC = ContentDetection.IsFlatOPCContent(val);

                        if (isFlatOPC)
                        {
                            // <?mso-application progid="Word.Document"?>
                            // <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">

                            log.Debug(".. contains Flat OPC content ");
                            cc.Type = Word.WdContentControlType.wdContentControlRichText;
                            // Ensure block level
                            Inline2Block i2b = new Inline2Block();
                            cc = i2b.convertToBlockLevel(cc, false, true);

                            if (cc == null)
                            {
                                MessageBox.Show("Problems inserting block level WordML at this location.");
                                return;
                            }
                            td.set("od:progid", "Word.Document");
                            cc.Title = "Word: " + xppe.xpathId;
                            //cc.Range.Text = val; // don't escape it
                            cc.Range.InsertXML(val, ref missing);

                        }
                        else if (ContentDetection.IsBase64Encoded(val))
                        {

                            // Force picture content control ...
                            // cc.Type = Word.WdContentControlType.wdContentControlPicture;
                            // from wdContentControlText (or wdContentControlRichText for that matter)
                            // doesn't work (you get "inappropriate range for applying this
                            // content control type").

                            // They've said map, so delete existing, and replace it.
                            designMode.Off();
                            cc.Delete(true);

                            // Now add a new cc
                            Globals.ThisAddIn.Application.Selection.Collapse(ref missing);
                            if (_PictureContentControlsReplace)
                            {
                                // Use a rich text control instead
                                cc = CurrentDocument.ContentControls.Add(
                                    Word.WdContentControlType.wdContentControlRichText, ref missing);

                                PictureUtils.setPictureHandler(td);
                                cc.Title = "Image: " + xppe.xpathId;

                                PictureUtils.pastePictureIntoCC(cc, Convert.FromBase64String(val));
                            }
                            else
                            {
                                cc = CurrentDocument.ContentControls.Add(
                                    Word.WdContentControlType.wdContentControlPicture, ref missing);
                            }

                            designMode.restoreState();

                        }
                        else
                        {
                            isXHTML = ContentDetection.IsXHTMLContent(val);
                        }

                        if (cc.Type == Word.WdContentControlType.wdContentControlText)
                        {
                            // cc.Type = Word.WdContentControlType.wdContentControlText;  // ???
                            cc.MultiLine = true;
                        }




                        //if (HasXHTMLContent(tn))
                        if (isXHTML)
                        {
                            log.Info("detected XHTML.. ");
                            td.set("od:ContentType", "application/xhtml+xml");
                            cc.Title = "XHTML: " + xppe.xpathId;
                            cc.Type = Word.WdContentControlType.wdContentControlRichText;

                            cc.Range.Text = val; // don't escape it

                            if (Inline2Block.containsBlockLevelContent(val))
                            {
                                Inline2Block i2b = new Inline2Block();
                                cc = i2b.convertToBlockLevel(cc, true, true);

                                if (cc == null)
                                {
                                    MessageBox.Show("Problems inserting block level XHTML at this location.");
                                    return;
                                }

                            }
                        }
                        else if (!isFlatOPC)
                        {
                            cc.Title = "Data value: " + xppe.xpathId;
                        }

                        cc.Tag = td.asQueryString();

                        if (cc.Type != Word.WdContentControlType.wdContentControlRichText)
                        {
                            cc.XMLMapping.SetMappingByNode(
                                Utilities.MxnFromTn(controlTreeView.treeView.SelectedNode, CurrentPart, true));
                        }

                    }
                    else if ((controlMain.modeControlEnabled == true && controlMain.controlMode1.isModeCondition())
                        || odType == ControlTreeView.OpenDopeType.Condition)
                    {
                        log.Debug("In condition mode");

                        // We want to be in Design Mode, so user can see their gesture take effect
                        designMode.On();

                        ConditionsPartEntry cpe = new ConditionsPartEntry(controlMain.model);
                        cpe.setup(CurrentPart.Id, strXPath, prefixMappings, true);
                        cpe.save();

                        cc.Title = "Conditional: " + cpe.conditionId;
                        // Write tag
                        td.set("od:condition", cpe.conditionId);
                        cc.Tag = td.asQueryString();

                        // Make it RichText; remove any pre-existing bind
                        if (cc.XMLMapping.IsMapped)
                        {
                            cc.XMLMapping.Delete();
                        }
                        if (cc.Type == Word.WdContentControlType.wdContentControlText)
                        {
                            cc.Type = Word.WdContentControlType.wdContentControlRichText;
                        }
                    }
                    else if ((controlMain.modeControlEnabled == true && controlMain.controlMode1.isModeRepeat())
                        || odType == ControlTreeView.OpenDopeType.Repeat)
                    {
                        log.Debug("In repeat mode");

                        // We want to be in Design Mode, so user can see their gesture take effect
                        designMode.On();

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

                        // Make it RichText; remove any pre-existing bind
                        if (cc.XMLMapping.IsMapped)
                        {
                            cc.XMLMapping.Delete();
                        }
                        if (cc.Type == Word.WdContentControlType.wdContentControlText)
                        {
                            cc.Type = Word.WdContentControlType.wdContentControlRichText;
                        }

                    }


                    //ensure it's checked
                    controlTreeView.mapToSelectedControlToolStripMenuItem.Checked = true;
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
