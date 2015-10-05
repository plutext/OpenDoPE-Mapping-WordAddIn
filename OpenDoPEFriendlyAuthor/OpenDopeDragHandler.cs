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
    class OpenDopeDragHandler
    {
        static Logger log = LogManager.GetLogger("OpenDopeDragHandler");

        public void treeView_ItemDrag(object sender, ItemDragEventArgs e,
            ControlTreeView controlTreeView,
            ControlMain controlMain, 
            Word.Document CurrentDocument, 
            Office.CustomXMLPart CurrentPart, XmlDocument OwnerDocument,
            bool _PictureContentControlsReplace)
        {
            object missing = System.Type.Missing;

            TreeNode tn = (TreeNode)e.Item;

            if (tn == null)
            {
                Debug.Fail("no tn");
                return;
            }

            //check if this is something we can drag
            if (((XmlNode)tn.Tag).NodeType == XmlNodeType.ProcessingInstruction
                || ((XmlNode)tn.Tag).NodeType == XmlNodeType.Comment)
                return;
            if (controlMain.modeControlEnabled == false // ie always mode bind
                || controlMain.controlMode1.isModeBind())
            {
                if (!ControlTreeView.IsLeafNode(tn)
                    || ((XmlNode)tn.Tag).NodeType == XmlNodeType.Text && !ControlTreeView.IsLeafNode(tn.Parent))
                    return;
            } // repeats and conditions; let them drag any node


            //get an nsmgr
            NameTable nt = new NameTable();

            //generate the xpath and the ns manager
            XmlNamespaceManager xmlnsMgr = new XmlNamespaceManager(nt);
            string strXPath = Utilities.XpathFromXn(CurrentPart.NamespaceManager, (XmlNode)tn.Tag, true, xmlnsMgr);
            log.Info("Dragging XPath: " + strXPath);

            string prefixMappings = Utilities.GetPrefixMappings(xmlnsMgr);

            // OpenDoPE
            TagData td = new TagData("");
            String val = ((XmlNode)tn.Tag).InnerText;

            DesignMode designMode = new OpenDoPEModel.DesignMode(CurrentDocument);

            // Special case for pictures, since drag/drop does not seem
            // to work properly (the XHTML pasted doesn't do what it should?)
            bool isPicture = ContentDetection.IsBase64Encoded(val);

            if (isPicture && !_PictureContentControlsReplace)
            {
                designMode.Off();
                log.Debug("Special case handling for pictures..");

                // Selection can't be textual content, so ensure it isn't.
                // It is allowed to be a picture, so in the future we could
                // leave the selection alone if it is just a picture.
                Globals.ThisAddIn.Application.Selection.Collapse(ref missing);
                // Are they dragging to an existing picture content control
                Word.ContentControl picCC = ContentControlMaker.getActiveContentControl(CurrentDocument, Globals.ThisAddIn.Application.Selection);
                try
                {
                    if (picCC == null
                        || (picCC.Type != Word.WdContentControlType.wdContentControlPicture))
                    {
                        picCC = CurrentDocument.ContentControls.Add(
                            Word.WdContentControlType.wdContentControlPicture, ref missing);
                        designMode.restoreState();
                    }
                }
                catch (COMException ce)
                {
                    // Will happen if you try to drag a text node onto an existing image content control
                    log.Debug("Ignoring " + ce.Message);
                    return;
                }
                XPathsPartEntry xppe = new XPathsPartEntry(controlMain.model);
                xppe.setup(null, CurrentPart.Id, strXPath, prefixMappings, false);
                xppe.save();

                td.set("od:xpath", xppe.xpathId);
                picCC.Tag = td.asQueryString();

                picCC.Title = "Data value: " + xppe.xpathId;
                picCC.XMLMapping.SetMappingByNode(Utilities.MxnFromTn(tn, CurrentPart, true));
                return;
            }

            log.Debug("\n\ntreeView_ItemDrag for WdSelectionType " + Globals.ThisAddIn.Application.Selection.Type.ToString());

            bool isXHTML = false;
            bool isFlatOPC = ContentDetection.IsFlatOPCContent(val);
            if (!isFlatOPC) isXHTML = ContentDetection.IsXHTMLContent(val);

            if (Globals.ThisAddIn.Application.Selection.Type
                != Microsoft.Office.Interop.Word.WdSelectionType.wdSelectionIP)
            {
                // ie something is selected, since "inline paragraph selection" 
                // just means the cursor is somewhere inside
                // a paragraph, but with nothing selected.

                designMode.Off();

                // Selection types: http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdselectiontype(v=office.11).aspx
                log.Debug("treeView_ItemDrag fired, but interpreted as gesture for WdSelectionType " + Globals.ThisAddIn.Application.Selection.Type.ToString());

                Word.ContentControl parentCC = ContentControlMaker.getActiveContentControl(CurrentDocument,
                            Globals.ThisAddIn.Application.Selection);

                // Insert bind | condition | repeat
                // depending on which mode button is pressed.
                if (controlMain.modeControlEnabled == false // ie always mode bind
                    || controlMain.controlMode1.isModeBind())
                {
                    log.Debug("In bind mode");
                    Word.ContentControl cc = null;


                    try
                    {
                        if (isFlatOPC || isXHTML
                            || (isPicture && _PictureContentControlsReplace))
                        {

                            // Rich text
                            if (parentCC != null
                                && ContentControlOpenDoPEType.isBound(parentCC))
                            {
                                // Reuse existing cc
                                cc = ContentControlMaker.MakeOrReuse(true, Word.WdContentControlType.wdContentControlRichText, CurrentDocument,
                                    Globals.ThisAddIn.Application.Selection);
                            }
                            else
                            {
                                // Make new cc
                                cc = ContentControlMaker.MakeOrReuse(true, Word.WdContentControlType.wdContentControlRichText, CurrentDocument,
                                    Globals.ThisAddIn.Application.Selection);
                            }

                            if (isFlatOPC)
                            {
                                log.Debug(".. contains block content ");
                                // Ensure block level
                                Inline2Block i2b = new Inline2Block();
                                cc = i2b.convertToBlockLevel(cc, false, true);

                                if (cc == null)
                                {
                                    MessageBox.Show("Problems inserting block level WordML at this location.");
                                    return;
                                }

                            }
                            else if (isXHTML // and thus not picture
                             && Inline2Block.containsBlockLevelContent(val))
                            {
                                log.Debug(".. contains block content ");
                                // Ensure block level
                                Inline2Block i2b = new Inline2Block();
                                cc = i2b.convertToBlockLevel(cc, false, true);

                                if (cc == null)
                                {
                                    MessageBox.Show("Problems inserting block level XHTML at this location.");
                                    return;
                                }
                            }

                        }
                        else
                        {

                            // Plain text
                            if (parentCC != null
                                && ContentControlOpenDoPEType.isBound(parentCC))
                            {
                                // Reuse existing cc
                                cc = ContentControlMaker.MakeOrReuse(true, Word.WdContentControlType.wdContentControlText, CurrentDocument,
                                    Globals.ThisAddIn.Application.Selection);
                            }
                            else
                            {
                                // Make new cc
                                cc = ContentControlMaker.MakeOrReuse(false, Word.WdContentControlType.wdContentControlText, CurrentDocument,
                                    Globals.ThisAddIn.Application.Selection);
                            }
                            cc.MultiLine = true;
                            // Is a text content control always run-level?
                            // No, not if you have a single para selected and you do drag gesture
                            // (or if you remap a rich text control)
                        }
                    }
                    catch (Exception ex)
                    {
                        log.Error("Couldn't add content control: " + ex.Message);
                        return;
                    }

                    XPathsPartEntry xppe = new XPathsPartEntry(controlMain.model);
                    xppe.setup(null, CurrentPart.Id, strXPath, prefixMappings, true);
                    xppe.save();

                    td.set("od:xpath", xppe.xpathId);

                    if (isFlatOPC)
                    {
                        // <?mso-application progid="Word.Document"?>
                        // <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">

                        td.set("od:progid", "Word.Document");
                        cc.Title = "Word: " + xppe.xpathId;
                        //cc.Range.Text = val; // don't escape it
                        cc.Range.InsertXML(val, ref missing);

                    }
                    else if (isXHTML)
                    {
                        td.set("od:ContentType", "application/xhtml+xml");
                        cc.Title = "XHTML: " + xppe.xpathId;
                        cc.Range.Text = val; // don't escape it

                    }
                    else if (isPicture)
                    {

                        PictureUtils.setPictureHandler(td);
                        cc.Title = "Image: " + xppe.xpathId;

                        string picContent = CurrentPart.SelectSingleNode(strXPath).Text;
                        PictureUtils.pastePictureIntoCC(cc, Convert.FromBase64String(picContent));

                    }
                    else
                    {
                        cc.XMLMapping.SetMappingByNode(Utilities.MxnFromTn(tn, CurrentPart, true));

                        string nodeXML = cc.XMLMapping.CustomXMLNode.XML;
                        log.Info(nodeXML);
                        cc.Title = "Data value: " + xppe.xpathId;
                    }

                    cc.Tag = td.asQueryString();

                    designMode.restoreState();

                }
                else if (controlMain.controlMode1.isModeCondition())
                {

                    log.Debug("In condition mode");
                    Word.ContentControl cc = null;
                    try
                    {
                        // always make new
                        cc = ContentControlMaker.MakeOrReuse(false, Word.WdContentControlType.wdContentControlRichText, CurrentDocument, Globals.ThisAddIn.Application.Selection);
                    }
                    catch (Exception ex)
                    {
                        log.Error("Couldn't add content control: " + ex.Message);
                        return;
                    }
                    ConditionsPartEntry cpe = new ConditionsPartEntry(controlMain.model);
                    cpe.setup(CurrentPart.Id, strXPath, prefixMappings, true);
                    cpe.save();

                    cc.Title = "Conditional: " + cpe.conditionId;
                    // Write tag
                    td.set("od:condition", cpe.conditionId);
                    cc.Tag = td.asQueryString();

                    designMode.On();

                }
                else if (controlMain.controlMode1.isModeRepeat())
                {
                    log.Debug("In repeat mode");
                    Word.ContentControl cc = null;
                    try
                    {
                        // always make new
                        cc = ContentControlMaker.MakeOrReuse(false, Word.WdContentControlType.wdContentControlRichText, CurrentDocument, Globals.ThisAddIn.Application.Selection);
                    }
                    catch (Exception ex)
                    {
                        log.Error("Couldn't add content control: " + ex.Message);
                        return;
                    }

                    // Need to drop eg [1] (if any), so BetterForm-based interactive processing works
                    if (strXPath.EndsWith("]"))
                    {
                        strXPath = strXPath.Substring(0, strXPath.LastIndexOf("["));
                        log.Debug("Having dropped '[]': " + strXPath);
                    }

                    XPathsPartEntry xppe = new XPathsPartEntry(controlMain.model);
                    xppe.setup("rpt", CurrentPart.Id, strXPath, prefixMappings, false); // no Q for repeat
                    xppe.save();

                    cc.Title = "Repeat: " + xppe.xpathId;
                    // Write tag
                    td.set("od:repeat", xppe.xpathId);
                    cc.Tag = td.asQueryString();

                    designMode.On();
                }

                return;
            } // end if (Globals.ThisAddIn.Application.Selection.Type != Microsoft.Office.Interop.Word.WdSelectionType.wdSelectionIP) 

            // Selection.Type: Microsoft.Office.Interop.Word.WdSelectionType.wdSelectionIP
            // ie cursor is somewhere inside a paragraph, but with nothing selected.
            log.Info("In wdSelectionIP specific code.");

            // leave designMode alone here

            // Following processing uses clipboard HTML to implement drag/drop processing
            // Could avoid dealing with that (what's the problem anyway?) if they are dragging onto an existing content control, with:
            //Word.ContentControl existingCC = ContentControlMaker.getActiveContentControl(CurrentDocument, Globals.ThisAddIn.Application.Selection);
            //if (existingCC != null) return;
            // But that stops them from dragging any more content into a repeat. 

            string title = "";
            string tag = "";
            bool needBind = false;

            log.Debug(strXPath);
            Office.CustomXMLNode targetNode = CurrentPart.SelectSingleNode(strXPath);
            string nodeContent = targetNode.Text;
            // or ((XmlNode)tn.Tag).InnerXml

            // Insert bind | condition | repeat
            // depending on which mode button is pressed.
            if (controlMain.modeControlEnabled == false // ie always mode bind
                || controlMain.controlMode1.isModeBind())
            {
                log.Debug("In bind mode");
                // OpenDoPE: create w:tag=od:xpath=x1
                // and add XPath to xpaths part
                XPathsPartEntry xppe = new XPathsPartEntry(controlMain.model);
                xppe.setup(null, CurrentPart.Id, strXPath, prefixMappings, false); // Don't setup Q until after drop
                xppe.save();

                // Write tag
                td.set("od:xpath", xppe.xpathId);

                // Does this node contain XHTML?
                // TODO: error handling
                log.Info(nodeContent);


                if (isFlatOPC)
                {
                    // <?mso-application progid="Word.Document"?>
                    // <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">

                    td.set("od:progid", "Word.Document");
                    title = "Word: " + xppe.xpathId;
                    needBind = false; // make it a rich text control

                }
                else if (isXHTML)
                {

                    td.set("od:ContentType", "application/xhtml+xml");
                    // TODO since this is a run-level sdt, 
                    // the XHTML content will need to be run-level.
                    // Help the user with this?
                    // Or in a run-level context, docx4j could convert
                    // p to soft-enter?  But what to do about tables?
                    title = "XHTML: " + xppe.xpathId;

                    needBind = false; // make it a rich text control

                    // Word will only replace our HTML-imported-to-docx with the raw HTML
                    // if we have the bind.
                    // Without this, giving the user visual feedback in Word is a TODO
                }
                else if (isPicture)
                {
                    designMode.Off();
                    log.Debug("NEW Special case handling for pictures..");

                    //object missing = System.Type.Missing;
                    Globals.ThisAddIn.Application.Selection.Collapse(ref missing);
                    // Are they dragging to an existing picture content control
                    Word.ContentControl picCC = ContentControlMaker.getActiveContentControl(CurrentDocument, Globals.ThisAddIn.Application.Selection);
                    try
                    {
                        if (picCC == null
                            || (picCC.Type != Word.WdContentControlType.wdContentControlPicture))
                        {
                            picCC = CurrentDocument.ContentControls.Add(
                                Word.WdContentControlType.wdContentControlRichText, ref missing);
                            designMode.restoreState();
                        }
                    }
                    catch (COMException ce)
                    {
                        // Will happen if you try to drag a text node onto an existing image content control
                        log.Debug("Ignoring " + ce.Message);
                        return;
                    }
                    PictureUtils.setPictureHandler(td);
                    picCC.Title = "Image: " + xppe.xpathId;

                    picCC.Tag = td.asQueryString();

                    PictureUtils.pastePictureIntoCC(picCC,
                        Convert.FromBase64String(nodeContent));

                    return;

                }
                else
                {
                    title = "Data value: " + xppe.xpathId;
                    needBind = true;
                }
                tag = td.asQueryString();

            }
            else if (controlMain.controlMode1.isModeCondition())
            {
                log.Debug("In condition mode");
                ConditionsPartEntry cpe = new ConditionsPartEntry(controlMain.model);
                cpe.setup(CurrentPart.Id, strXPath, prefixMappings, false);
                cpe.save();

                title = "Conditional: " + cpe.conditionId;
                // Write tag
                td.set("od:condition", cpe.conditionId);
                tag = td.asQueryString();
            }
            else if (controlMain.controlMode1.isModeRepeat())
            {
                log.Debug("In repeat mode");

                // Need to drop eg [1] (if any), so BetterForm-based interactive processing works
                if (strXPath.EndsWith("]"))
                {
                    strXPath = strXPath.Substring(0, strXPath.LastIndexOf("["));
                    log.Debug("Having dropped '[]': " + strXPath);
                }

                XPathsPartEntry xppe = new XPathsPartEntry(controlMain.model);
                xppe.setup("rpt", CurrentPart.Id, strXPath, prefixMappings, false);
                xppe.save();

                title = "Data value: " + xppe.xpathId;
                // Write tag
                td.set("od:repeat", xppe.xpathId);
                tag = td.asQueryString();
            }


            //create the HTML
            string strHTML = string.Empty;
            if (isFlatOPC)
            {
                // <?mso-application progid="Word.Document"?>
                // <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
                nodeContent = ControlTreeView.EscapeXHTML(nodeContent);

                strHTML = ClipboardUtilities.GenerateClipboardHTML(needBind, strXPath, Utilities.GetPrefixMappings(xmlnsMgr), CurrentPart.Id,
                    Utilities.MappingType.RichText, title, tag, nodeContent);

            }
            else if (isXHTML)
            {
                // need to escape eg <span> for it to get through the Clipboard
                nodeContent = ControlTreeView.EscapeXHTML(nodeContent);

                // use a RichText control, and set nodeContent
                strHTML = ClipboardUtilities.GenerateClipboardHTML(needBind, strXPath, Utilities.GetPrefixMappings(xmlnsMgr), CurrentPart.Id,
                    Utilities.MappingType.RichText, title, tag, nodeContent);
                // alternatively, this could be done in DocumentEvents.doc_ContentControlAfterAdd
                // but to do it there, we'd need to manually resolve the XPath to
                // find the value of the CustomXMLNode it pointed to.

            }
            else if (!needBind)
            {
                // For conditions & repeats, we use a RichText control
                strHTML = ClipboardUtilities.GenerateClipboardHTML(needBind, strXPath, Utilities.GetPrefixMappings(xmlnsMgr), CurrentPart.Id,
                    Utilities.MappingType.RichText, title, tag);
            }
            else
            {
                // Normal bind

                if (OwnerDocument.Schemas.Count > 0)
                {
                    switch (Utilities.CheckNodeType((XmlNode)tn.Tag))
                    {
                        case Utilities.MappingType.Date:
                            strHTML = ClipboardUtilities.GenerateClipboardHTML(needBind, strXPath, prefixMappings, CurrentPart.Id, Utilities.MappingType.Date, title, tag);
                            break;
                        case Utilities.MappingType.DropDown:
                            strHTML = ClipboardUtilities.GenerateClipboardHTML(needBind, strXPath, prefixMappings, CurrentPart.Id, Utilities.MappingType.DropDown, title, tag);
                            break;
                        case Utilities.MappingType.Picture:
                            strHTML = ClipboardUtilities.GenerateClipboardHTML(needBind, strXPath, prefixMappings, CurrentPart.Id, Utilities.MappingType.Picture, title, tag);
                            break;
                        default:
                            strHTML = ClipboardUtilities.GenerateClipboardHTML(needBind, strXPath, prefixMappings, CurrentPart.Id, Utilities.MappingType.Text, title, tag);
                            break;
                    }
                }
                else
                {
                    //String val = ((XmlNode)tn.Tag).InnerText;
                    if (ContentDetection.IsBase64Encoded(val))
                    {
                        // Force picture content control
                        strHTML = ClipboardUtilities.GenerateClipboardHTML(needBind, strXPath, Utilities.GetPrefixMappings(xmlnsMgr), CurrentPart.Id,
                            Utilities.MappingType.Picture, title, tag);
                    }
                    else
                    {
                        strHTML = ClipboardUtilities.GenerateClipboardHTML(needBind, strXPath, Utilities.GetPrefixMappings(xmlnsMgr), CurrentPart.Id,
                            Utilities.MappingType.Text, title, tag);
                    }
                }
            }

            // All cases:-

            //notify ourselves of a pending drag/drop
            controlMain.NotifyDragDrop(true);

            //throw it on the clipboard to drag
            DataObject dobj = new DataObject();
            dobj.SetData(DataFormats.Html, strHTML);
            dobj.SetData(DataFormats.Text, tn.Text);
            controlTreeView.DoDragDrop(dobj, DragDropEffects.Move);

            //notify ourselves of a completed drag/drop
            controlMain.NotifyDragDrop(false);

            Clipboard.SetData(DataFormats.Text, ((XmlNode)tn.Tag).InnerXml);
        }
    }
}
