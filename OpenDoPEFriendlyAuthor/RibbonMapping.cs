/*
 * (c) Copyright Plutext Pty Ltd, 2012
 * 
 * All rights reserved.
 * 
 * This source code is the proprietary information of Plutext
 * Pty Ltd, and must be kept confidential.
 * 
 * You may use, modify and distribute this source code only
 * as provided in your license agreement with Plutext.
 * 
 * If you do not have a license agreement with Plutext:
 * 
 * (i) you must return all copies of this source code to Plutext, 
 * or destroy it.  
 * 
 * (ii) under no circumstances may you use, modify or distribute 
 * this source code.
 * 
 *
 * Portions Copyright (c) Microsoft Corporation.  All rights reserved.
 * 
 * With respect to those portions (from http://xmlmapping.codeplex.com/license):
 * 

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
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using NLog;
using OpenDoPEModel;
using System.Collections.Generic;

namespace XmlMappingTaskPane
{
    public partial class RibbonMapping
    {
        // Note, there is no event which allows us to detect
        // that the user has changed to this ribbon tab.
        // May be able to do it using accessibility?

        static Logger log = LogManager.GetLogger("RibbonMapping");

        private void RibbonMapping_Load(object sender, RibbonUIEventArgs e)
        {
            //log.Debug("You clicked the ribbon");
            //CustomTaskPane ctpPaneForThisWindow = Utilities.FindTaskPaneForCurrentWindow();
            //// If the task pane exists, show it & grey out the import button
            //if (ctpPaneForThisWindow != null) {
            //    //it's built and being clicked, show it
            //    ctpPaneForThisWindow.Visible = true;
            //}

            //// else if the OpenDoPE parts are present in this docx, 
            //// launch the task pane 
            //if (CustomXmlUtilities.areOpenDoPEPartsPresent(Globals.ThisAddIn.Application.ActiveDocument))
            //{
            //    log.Info("OpenDoPE parts detected as present.. launching task pane");
            //    launchTaskPane();
            //}

            //// else, do nothing .. wait for user to choose to press
            //// the import button

            this.buttonBind.Enabled = false;
            this.buttonCondition.Enabled = false;
            this.buttonRepeat.Enabled = false;

            this.buttonEdit.Enabled = false;
            this.buttonDelete.Enabled = false;

            this.menuAdvanced.Enabled = false;

            this.buttonReplaceXML.Enabled = false;
            this.buttonClearAll.Enabled = false;
        }

        private void toggleButtonMapping_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;

            //get the ctp for this window (or null if there's not one)
            CustomTaskPane ctpPaneForThisWindow = Utilities.FindTaskPaneForCurrentWindow();


            if (toggleButtonMapping.Checked == false)
            {
                Debug.Assert(ctpPaneForThisWindow != null, 
                    "how was the ribbon button pressed if there was a control?");

                //it's being unclicked
                if (ctpPaneForThisWindow != null)
                    ctpPaneForThisWindow.Visible = false;

            }
            else if (ctpPaneForThisWindow == null) 
            {
                if (CustomXmlUtilities.areOpenDoPEPartsPresent(Globals.ThisAddIn.Application.ActiveDocument))
                {
                    log.Debug("OpenDoPE parts detected as present");
                }
                else
                {
                    log.Info("OpenDoPE parts not detected; adding now.");

                    Model model = Model.ModelFactory(document);
                    string requiredRoot = System.Configuration.ConfigurationManager.AppSettings["RootElement"];
                    Office.CustomXMLPart userPart = model.getUserPart(requiredRoot);

                    bool cancelled = false;
                    while (userPart == null && !cancelled)
                    {
                        using (Forms.FormAddPart fap = new Forms.FormAddPart())
                        {
                            //add a new stream from the XML retrieved from the Add New dialog
                            //otherwise, select the last selected item and populate with its xml
                            if (fap.ShowDialog() == DialogResult.OK)
                            {
                                object missing = System.Reflection.Missing.Value;
                                Office.CustomXMLPart newCxp = document.CustomXMLParts.Add(fap.XmlString, missing);

                                model = Model.ModelFactory(document);
                                userPart = model.getUserPart(requiredRoot);
                                if (userPart == null)
                                {
                                    newCxp.Delete();
                                    MessageBox.Show("You need to use root element: " + requiredRoot);
                                }
                            }
                            else
                            {
                                cancelled = true;
                            }
                        }
                    }
                    if (cancelled)
                    {
                        MessageBox.Show("For template authoring, you need to add your XML.");
                        toggleButtonMapping.Checked = false;
                        return;
                    }

                    InitialSetup init = new InitialSetup();
                    init.process();
                }
                launchTaskPane();
            }
            else
            {
                //it's built and being clicked, show it
                ctpPaneForThisWindow.Visible = true;
            }
            
        }

        /// <summary>
        /// Replace the contents of the existing XML part.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void buttonReplaceXML_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            CustomTaskPane ctpPaneForThisWindow = Utilities.FindTaskPaneForCurrentWindow();
            Controls.ControlMain ccm = (Controls.ControlMain)ctpPaneForThisWindow.Control;

            Office.CustomXMLPart userPart = null;
            bool cancelled = false;
            while (userPart == null && !cancelled)
            {
                using (Forms.FormAddPart fap = new Forms.FormAddPart())
                {
                    if (fap.ShowDialog() == DialogResult.OK)
                    {

                        string requiredRoot = System.Configuration.ConfigurationManager.AppSettings["RootElement"];
                        if (string.IsNullOrWhiteSpace(requiredRoot))
                        {
                            ccm.EventHandlerAndOnChildren.NodeAfterReplaceDisconnect();
                            Office.CustomXMLPart existingPart = ccm.model.getUserPart(requiredRoot);
                            CustomXmlUtilities.replaceXmlDoc(existingPart, fap.XmlString);
                            userPart = existingPart;
                        }
                        else
                        {
                            Office.CustomXMLPart existingPart = ccm.model.getUserPart(requiredRoot);
                            if (existingPart.DocumentElement.BaseName.Equals(requiredRoot))
                            {
                                ccm.EventHandlerAndOnChildren.NodeAfterReplaceDisconnect();
                                CustomXmlUtilities.replaceXmlDoc(existingPart, fap.XmlString);
                                userPart = existingPart;
                            }
                            else
                            {
                                MessageBox.Show("You need to use root element: " + requiredRoot);
                            }
                        }
                    }
                    else
                    {
                        cancelled = true;
                    }

                }
            }

            if (!cancelled)
            {
                ccm.EventHandlerAndOnChildren.NodeAfterReplaceReconnect();
            }
            ccm.RefreshTreeControl(null);

        }

        void buttonClearAll_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            // First, are you sure?
            DialogResult result = MessageBox.Show("This will remove your content controls (but keep their contents), and the XML they were mapped to.", "Are you sure?", MessageBoxButtons.OKCancel);
            if (!result.Equals(DialogResult.OK)) return;

            // Remove all content controls, but keep contents
            Word.Document docx = Globals.ThisAddIn.Application.ActiveDocument;
            foreach (Word.Range storyRange in docx.StoryRanges)
            {
                foreach (Word.ContentControl cc in storyRange.ContentControls)
                {
                    // Delete the content control, but keep the contents
                    cc.Delete(false);
                }
            }

            // That's not enough to get the headers/footers!
            foreach (Word.Section section in docx.Sections)
            {
                foreach (Word.HeaderFooter hf in section.Headers) {
                    foreach (Word.ContentControl cc in hf.Range.ContentControls) {
                        cc.Delete(false);
                    }
                }
                foreach (Word.HeaderFooter hf in section.Footers)
                {
                    foreach (Word.ContentControl cc in hf.Range.ContentControls)
                    {
                        cc.Delete(false);
                    }
                }
            }


                
            CustomTaskPane ctpPaneForThisWindow = Utilities.FindTaskPaneForCurrentWindow();
            if (ctpPaneForThisWindow != null)
            {
                Controls.ControlMain ccm = (Controls.ControlMain)ctpPaneForThisWindow.Control;
                ccm.EventHandlerAndOnChildren.DisconnectOldDocument();

                // Remove all Custom XML Parts
                ccm.model.RemoveParts();
                //Model model = Model.ModelFactory(docx);
                //model.RemoveParts();

                // Remove task pane
                Globals.ThisAddIn.TaskPaneList.Remove(Globals.ThisAddIn.Application.ActiveWindow);
                Globals.ThisAddIn.CustomTaskPanes.Remove(ctpPaneForThisWindow);
                ctpPaneForThisWindow.Dispose();
            }

            // Grey out buttons
            this.buttonBind.Enabled = false;
            this.buttonCondition.Enabled = false;
            this.buttonRepeat.Enabled = false;

            this.buttonEdit.Enabled = false;
            this.buttonDelete.Enabled = false;

            this.menuAdvanced.Enabled = false;

            this.buttonReplaceXML.Enabled = false;
            this.buttonClearAll.Enabled = false;

            toggleButtonMapping.Checked = false;


            // Tell the user
            MessageBox.Show("Content controls and XML parts have been deleted.", "Deletion Confirmed", MessageBoxButtons.OK);
        }


        private void launchTaskPane()
        {
            //set the cursor to wait
            Globals.ThisAddIn.Application.System.Cursor = Word.WdCursorType.wdCursorWait;

            //set up the task pane
            CustomTaskPane ctpPaneForThisWindow 
                = Globals.ThisAddIn.CustomTaskPanes.Add(
                    new Controls.ControlMain(), 
                    Properties.Resources.TaskPaneName);
            ctpPaneForThisWindow.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal;

            //store it for later
            if (Globals.ThisAddIn.Application.ShowWindowsInTaskbar)
                Globals.ThisAddIn.TaskPaneList.Add(Globals.ThisAddIn.Application.ActiveWindow, ctpPaneForThisWindow);

            //connect task pane events
            Globals.ThisAddIn.ConnectTaskPaneEvents(ctpPaneForThisWindow);

            //get the control we hosted
            Controls.ControlMain ccm = (Controls.ControlMain)ctpPaneForThisWindow.Control;

            ccm.formPartList.ccm = ccm;

            // OK, OpenDoPE parts should exist
            Model model = Model.ModelFactory(Globals.ThisAddIn.Application.ActiveDocument);
            // hand the OpenDoPE model to the control
            ccm.model = model;


            //hand the eventing class to the control
            DocumentEvents de = new DocumentEvents(ccm);
            de.RibbonMapping = this; // so de can enable/disable stuff on ribbon


            ccm.EventHandlerAndOnChildren = de;

            // following is no good for us, since it selects first part
            //ccm.RefreshControls(Controls.ControlMain.ChangeReason.DocumentChanged, null, null, null, null, null);

            // The current cxp is stored in the DocumentEvents object
            // (which in turn is a field in ControlBase).
            // The user custom xml part is to current in DocumentEvents.
            // But we also need:
            string requiredRoot = System.Configuration.ConfigurationManager.AppSettings["RootElement"];
            Office.CustomXMLPart userPart = model.getUserPart(requiredRoot);
            Office.CustomXMLNode mxnNewNode = userPart.DocumentElement;

            log.Debug("DocumentElement of user part: " + mxnNewNode.BaseName);
            ccm.formPartList.controlPartList.RefreshPartList(true, false, false, mxnNewNode.OwnerPart.Id, null, mxnNewNode);

            //ccm.RefreshControls(Controls.ControlMain.ChangeReason.OnEnter, null, null, null, mxnNewNode, null);
            // would also work, but is more opaque

            //show it                            
            ctpPaneForThisWindow.Visible = true;

            this.buttonBind.Enabled = true;
            this.buttonCondition.Enabled = true;
            this.buttonRepeat.Enabled = true;

            this.menuAdvanced.Enabled = true;

            this.buttonReplaceXML.Enabled = true;
            this.buttonClearAll.Enabled = true;

            //reset the cursor
            Globals.ThisAddIn.Application.System.Cursor = Word.WdCursorType.wdCursorNormal;

        }

        private void buttonBind_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;

            // Workaround for reported Word crash.
            // Can't reproduce with Word 2010 sp1: 14.0.6129.500
            Word.ContentControl currentCC = ContentControlMaker.getActiveContentControl(document, Globals.ThisAddIn.Application.Selection);
            if (currentCC != null && currentCC.Type != Word.WdContentControlType.wdContentControlRichText)
            {
                MessageBox.Show("You can't add a data value here.");
                return;
            }

            OpenDoPEModel.DesignMode designMode = new OpenDoPEModel.DesignMode(document);
            designMode.Off();

            Word.ContentControl cc = null;
            object missing = System.Type.Missing;
            try
            {
                cc = document.ContentControls.Add(Word.WdContentControlType.wdContentControlText, ref missing);
                // Later, we'll change it to type picture if necessary
                cc.MultiLine = true;

                //                OpenDoPEModel.ContentControlStyle.adapt(cc);
            }
            catch (System.Exception)
            {
                MessageBox.Show("Selection must be either part of a single paragraph, or one or more whole paragraphs");
                return;
            }
            finally
            {
                designMode.restoreState();
            }

            cc.Title = "Data Value [unbound]"; // // This used if they later click edit

            // Now just launch the edit button
            editXPath(cc);
        }

        private void buttonCondition_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;

            // Workaround for reported Word crash.
            // Can't reproduce with Word 2010 sp1: 14.0.6129.500
            Word.ContentControl currentCC = ContentControlMaker.getActiveContentControl(document, Globals.ThisAddIn.Application.Selection);
            if (currentCC != null && currentCC.Type != Word.WdContentControlType.wdContentControlRichText)
            {
                MessageBox.Show("You can't add a condition here.");
                return;
            }

            OpenDoPEModel.DesignMode designMode = new OpenDoPEModel.DesignMode(document);
            designMode.Off();

            // Find a content control within the selection
            List<Word.ContentControl> shallowChildren = ContentControlUtilities.getShallowestSelectedContentControls(document, Globals.ThisAddIn.Application.Selection);
            log.Debug(shallowChildren.Count + " shallowChildren found.");

            Word.ContentControl conditionCC = null;
            object missing = System.Type.Missing;
            try
            {
                if (Globals.ThisAddIn.Application.Selection.Type == Microsoft.Office.Interop.Word.WdSelectionType.wdSelectionIP)
                {
                    // Nothing is selected, so type "condition"  
                    document.Windows[1].Selection.Text="condition";

                }
                object range = Globals.ThisAddIn.Application.Selection.Range;
                conditionCC = document.ContentControls.Add(Word.WdContentControlType.wdContentControlRichText, ref range);

                // Limitation here: you can't make your content control of eg type picture
                designMode.On();
            }
            catch (System.Exception)
            {
                MessageBox.Show("Selection must be either part of a single paragraph, or one or more whole paragraphs");
                designMode.restoreState();
                return;
            }

            conditionCC.Title = "Condition [unbound]"; // // This used if they later click edit

            if (shallowChildren.Count == 0)
            {
                log.Debug("No child control found. So Condition not set on our new CC");
                //MessageBox.Show("Unbound content control only added. Click the edit button to setup the condition.");
                editXPath(conditionCC);
                return;
            }
            // For now, just use the tag on the first simple bind we find. 
            // Later, we could try parsing a condition or repeat
            Word.ContentControl usableChild = null;
            foreach (Word.ContentControl child in shallowChildren)
            {
                //if (child.Tag.Contains("od:xpath"))
                if (child.XMLMapping.IsMapped)
                {
                    usableChild = child;
                    break;
                }
            }
            if (usableChild == null)
            {
                log.Debug("No usable child found. So Condition not set on our new CC");
                //MessageBox.Show("Naked content control only added. Click the edit button to setup the condition.");
                editXPath(conditionCC);
                return;
            }

            // Get XPath. Could use the od xpaths part, but 
            // easier here to get it from the binding
            string strXPath = usableChild.XMLMapping.XPath;
            log.Debug("Getting count condition from " + strXPath);
            strXPath = "count(" + strXPath + ")>0";
            log.Debug(strXPath);

            ConditionsPartEntry cpe = new ConditionsPartEntry(Model.ModelFactory(document));
            // TODO fix usableChild.XMLMapping.PrefixMappings
            cpe.setup(usableChild.XMLMapping.CustomXMLPart.Id, strXPath, "", true);
            cpe.save();

            conditionCC.Title = "Conditional: " + cpe.conditionId;
            // Write tag
            TagData td = new TagData("");
            td.set("od:condition", cpe.conditionId);
            conditionCC.Tag = td.asQueryString();

            editXPath(conditionCC);
        }

        // We can't add a tag for a repeat or a condition,
        // unless we know the XPath.  And we can't know that
        // if there is no suitable child to deduce it from.
        // But if they later press the edit button, we 
        // would prefer not to have to ask them whether it
        // is a repeat or a condition. We could track the
        // content controls by their ID, but it seems 
        // better to just use the Title.

        /// <summary>
        /// Wrap the selection in a Repeat.
        /// The XPath is deduced from plain binds in the contents.  
        /// If there are none of these, just insert an empty content control.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonRepeat_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;

            // Workaround for reported Word crash.
            // Can't reproduce with Word 2010 sp1: 14.0.6129.500
            Word.ContentControl currentCC = ContentControlMaker.getActiveContentControl(document, Globals.ThisAddIn.Application.Selection);
            if (currentCC != null && currentCC.Type != Word.WdContentControlType.wdContentControlRichText)
            {
                MessageBox.Show("You can't add a repeat here.");
                return;
            }

            OpenDoPEModel.DesignMode designMode = new OpenDoPEModel.DesignMode(document);
            designMode.Off();

            // Find a content control within the selection
            List<Word.ContentControl> shallowChildren = ContentControlUtilities.getShallowestSelectedContentControls(document, Globals.ThisAddIn.Application.Selection);
            log.Debug(shallowChildren.Count + " shallowChildren found.");

            Word.ContentControl repeatCC = null;
            object missing = System.Type.Missing;
            try
            {
                repeatCC = document.ContentControls.Add(Word.WdContentControlType.wdContentControlRichText, ref missing);
                // Limitation here: you can't make your content control of eg type picture
                log.Debug("New content control added, with id: " + repeatCC.ID);

                designMode.On();

            }
            catch (System.Exception) {
                MessageBox.Show("Selection must be either part of a single paragraph, or one or more whole paragraphs");
                designMode.restoreState();
                return; 
            }

            repeatCC.Title = "Repeat [unbound]"; // This used if they later click edit

            if (shallowChildren.Count == 0)
            {
                log.Debug("No child control found. So Repeat not set on our new CC");
                //MessageBox.Show("Unbound content control only added. Click the edit button to setup the repeat.");
                editXPath(repeatCC);
                return;
            }
            // For now, just use the tag on the first simple bind we find. 
            // Later, we could try parsing a condition or repeat
            Word.ContentControl usableChild = null;
            foreach (Word.ContentControl child in shallowChildren)
            {
                //if (child.Tag.Contains("od:xpath"))
                if (child.XMLMapping.IsMapped)
                {
                    usableChild = child;
                    break;
                }
            }
            if (usableChild == null)
            {
                log.Debug("No usable child found. So Repeat not set on our new CC");
                //MessageBox.Show("Naked content control only added. Click the edit button to setup the repeat.");
                editXPath(repeatCC);
                return;
            }

            string strXPath = null;             

            // Need to work out what repeats.  Could be this child,
            // the parent, etc.  If its obvious from this exemplar xml doc,
            // we can do it automatically. Otherwise, we'll ask user.
            // Currently the logic supports repeating ., .., or grandparent.

            // See whether this child repeats in this exemplar.
            Office.CustomXMLNode thisNode = usableChild.XMLMapping.CustomXMLNode;
            Office.CustomXMLNode thisNodeSibling = usableChild.XMLMapping.CustomXMLNode.NextSibling;
            Office.CustomXMLNode parent = usableChild.XMLMapping.CustomXMLNode.ParentNode;
            Office.CustomXMLNode parentSibling = usableChild.XMLMapping.CustomXMLNode.ParentNode.NextSibling;
            if (thisNodeSibling!=null
                && thisNodeSibling.BaseName.Equals(thisNode.BaseName) ) {
                // Looks like this node repeats :-)

                    strXPath = usableChild.XMLMapping.XPath;
                        // Get XPath. Could use the od xpaths part, but 
                        // easier here to work with the binding
                    log.Debug("Using . as repeat: " + strXPath);

            } // If it doesn't, test parent.
            else if (parentSibling != null
                && parentSibling.BaseName.Equals(parent.BaseName))
            {
                strXPath = usableChild.XMLMapping.XPath;
                log.Debug("Using parent for repeat ");
                strXPath = strXPath.Substring(0, strXPath.LastIndexOf("/"));
                log.Debug("Using: " + strXPath);
            }
            else // If that doesn't either, ask user. 
            {
                Office.CustomXMLNode grandparent = null;
                if (parent != null)
                {
                    grandparent = parent.ParentNode;
                }

                using (Forms.FormSelectRepeatedElement sr = new Forms.FormSelectRepeatedElement())
                {
                    sr.labelXPath.Text = usableChild.XMLMapping.XPath;
                    sr.listElementNames.Items.Add(thisNode.BaseName);
                    sr.listElementNames.Items.Add(parent.BaseName);
                    if (grandparent != null)
                    {
                        sr.listElementNames.Items.Add(grandparent.BaseName);
                    }
                    sr.listElementNames.SelectedIndex = 0;
                    sr.ShowDialog();
                    if (sr.listElementNames.SelectedIndex == 0)
                    {
                        strXPath = usableChild.XMLMapping.XPath;
                    }
                    else if (sr.listElementNames.SelectedIndex == 1)
                    {
                        strXPath = parent.XPath;
                        log.Debug("Using parent for repeat: " + strXPath);
                    }
                    else
                    {
                        // Grandparent
                        strXPath = grandparent.XPath;
                        log.Debug("Using grandparent for repeat: " + strXPath);
                    }
                }
            }

            // Need to drop eg [1] (if any), so BetterForm-based interactive processing works
            if (strXPath.EndsWith("]")) {
                strXPath = strXPath.Substring(0, strXPath.LastIndexOf("[")); 
                log.Debug("Having dropped '[]': " + strXPath);
            }

            XPathsPartEntry xppe = new XPathsPartEntry(Model.ModelFactory(document));
            // TODO fix usableChild.XMLMapping.PrefixMappings
            xppe.setup("rpt", usableChild.XMLMapping.CustomXMLPart.Id, strXPath, "", false);  // No Q for repeat
            xppe.save();

            repeatCC.Title = "Repeat: " + xppe.xpathId;
            // Write tag
            TagData td = new TagData("");
            td.set("od:repeat", xppe.xpathId);
            repeatCC.Tag = td.asQueryString();

        }


        private void buttonEdit_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;

            // Get the content control
            Word.ContentControl cc = ContentControlMaker.getActiveContentControl(document, Globals.ThisAddIn.Application.Selection);
            if (cc == null)
            {
                // Shouldn't happen
                MessageBox.Show("Which content control?");
                return;
            }
            editXPath(cc);
        }

        public static void editXPath(Word.ContentControl cc)
        {
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;

            // First, work out whether this is a condition or a repeat or a plain bind
            bool isCondition = false;
            bool isRepeat = false;
            bool isBind = false;
            if ( (cc.Title!=null && cc.Title.StartsWith("Condition") )
                || (cc.Tag!=null && cc.Tag.Contains("od:condition") ))
            {
                isCondition = true;
            }
            else if ( (cc.Title!=null && cc.Title.StartsWith("Repeat"))
                || (cc.Tag!=null && cc.Tag.Contains("od:repeat") ))
            {
                isRepeat = true;
            }
            else if ((cc.Title != null && cc.Title.StartsWith("Data"))
                || (cc.Tag != null && cc.Tag.Contains("od:xpath"))
                || cc.XMLMapping.IsMapped
                )
            {
                isBind = true;
            }
            else
            {
                // Ask user
                using (Forms.ConditionOrRepeat cor = new Forms.ConditionOrRepeat())
                {
                    if (cor.ShowDialog() == DialogResult.OK)
                    {
                        isCondition = cor.radioButtonCondition.Checked;
                        isRepeat = cor.radioButtonRepeat.Checked;
                        isBind = cor.radioButtonBind.Checked;
                    }
                    else
                    {
                        // They cancelled
                        return;
                    }
                }
            }

            // OK, now we know whether its a condition or a repeat or a bind
            // Is it already mapped to something?
            TagData td = new TagData(cc.Tag);
            Model model = Model.ModelFactory(document);

            string strXPath = "";

            // In order to get Id and prefix mappings for current part
            CustomTaskPane ctpPaneForThisWindow = Utilities.FindTaskPaneForCurrentWindow();
            Controls.ControlMain ccm = (Controls.ControlMain)ctpPaneForThisWindow.Control;

            string cxpId = ccm.CurrentPart.Id;
            string prefixMappings = ""; // TODO GetPrefixMappings(ccm.CurrentPart.NamespaceManager);
            log.Debug("default prefixMappings: " + prefixMappings);

            XPathsPartEntry xppe = null;
            ConditionsPartEntry cpe = null;

            if (isCondition
                && td.get("od:condition") != null)
            {
                string conditionId = td.get("od:condition");
                cpe = new ConditionsPartEntry(model);
                condition c = cpe.getConditionByID(conditionId);

                string xpathid = null;
                if (c!=null 
                    && c.Item is xpathref)
                {
                    xpathref ex = (xpathref)c.Item;
                    xpathid = ex.id;

                    // Now fetch the XPath
                    xppe = new XPathsPartEntry(model);

                    xpathsXpath xx = xppe.getXPathByID(xpathid);

                    if (xx != null)
                    {
                        strXPath = xx.dataBinding.xpath;
                        cxpId = xx.dataBinding.storeItemID;
                        prefixMappings = xx.dataBinding.prefixMappings;
                    }
                }
            }
            else if (isRepeat
              && td.get("od:repeat") != null)
            {
                string repeatId = td.get("od:repeat");

                // Now fetch the XPath
                xppe = new XPathsPartEntry(model);

                xpathsXpath xx = xppe.getXPathByID(repeatId);

                if (xx != null)
                {
                    strXPath = xx.dataBinding.xpath;
                    cxpId = xx.dataBinding.storeItemID;
                    prefixMappings = xx.dataBinding.prefixMappings;
                }
            }
            else if (isBind) {

              if (cc.XMLMapping.IsMapped) {
                // Prefer this, if for some reason it contradicts od:xpath
                strXPath = cc.XMLMapping.XPath;
                cxpId = cc.XMLMapping.CustomXMLPart.Id;
                prefixMappings = cc.XMLMapping.PrefixMappings;

              } else if( td.get("od:xpath") != null) {
                string xpathId = td.get("od:xpath");

                // Now fetch the XPath
                xppe = new XPathsPartEntry(model);

                xpathsXpath xx = xppe.getXPathByID(xpathId);

                if (xx != null)
                {
                    strXPath = xx.dataBinding.xpath;
                    cxpId = xx.dataBinding.storeItemID;
                    prefixMappings = xx.dataBinding.prefixMappings;
                }

              }
            }

            // Now we can present the form
            using (Forms.XPathEditor xpe = new Forms.XPathEditor())
            {
                xpe.textBox1.Text = strXPath;
                if (xpe.ShowDialog() == DialogResult.OK)
                {
                    strXPath = xpe.textBox1.Text;
                }
                else
                {
                    // They cancelled
                    return;
                }
            }

            // Now give effect to it
            td = new TagData("");
            if (isCondition)
            {
                // Create the new condition. Doesn't attempt to delete
                // the old one (if any)
                if (cpe == null)
                {
                    cpe = new ConditionsPartEntry(model);
                }
                cpe.setup(cxpId, strXPath, prefixMappings, true);
                cpe.save();

                cc.Title = "Conditional: " + cpe.conditionId;
                // Write tag
                td.set("od:condition", cpe.conditionId);
                cc.Tag = td.asQueryString();

            }
            else if (isRepeat)
            {
                // Create the new repeat. Doesn't attempt to delete
                // the old one (if any)
                if (xppe == null)
                {
                    xppe = new XPathsPartEntry(model);
                }

                xppe.setup("rpt", cxpId, strXPath, prefixMappings, false);
                xppe.save();

                cc.Title = "Repeat: " + xppe.xpathId;
                // Write tag
                td.set("od:repeat", xppe.xpathId);
                cc.Tag = td.asQueryString();
            }
            else if (isBind)
            {
                // Create the new bind. Doesn't attempt to delete
                // the old one (if any)
                if (xppe == null)
                {
                    xppe = new XPathsPartEntry(model);
                }

                Word.XMLMapping bind = cc.XMLMapping;
                bool mappable = bind.SetMapping(strXPath, prefixMappings, 
                    CustomXmlUtilities.getPartById(document, cxpId) );
                if (mappable) {
                    // What does the XPath point to?
                    string val = cc.XMLMapping.CustomXMLNode.Text;

                    cc.Title = "Data value: " + xppe.xpathId;

                    if (ContentDetection.IsBase64Encoded(val))
                    {
                        // Force picture content control ...
                        // cc.Type = Word.WdContentControlType.wdContentControlPicture;
                        // from wdContentControlText (or wdContentControlRichText for that matter)
                        // doesn't work (you get "inappropriate range for applying this
                        // content control type").

                        cc.Delete(true);

                        // Now add a new cc
                        object missing = System.Type.Missing;
                        Globals.ThisAddIn.Application.Selection.Collapse(ref missing);
                        cc = document.ContentControls.Add(
                            Word.WdContentControlType.wdContentControlPicture, ref missing);

                        cc.XMLMapping.SetMapping(strXPath, prefixMappings,
                            CustomXmlUtilities.getPartById(document, cxpId));

                    } else if (ContentDetection.IsXHTMLContent(val) )
                    {
                        td.set("od:ContentType", "application/xhtml+xml");
                        cc.Tag = td.asQueryString();

                        cc.XMLMapping.Delete();
                        cc.Type = Word.WdContentControlType.wdContentControlRichText;
                        cc.Title = "XHTML: " + xppe.xpathId;

                        if (Inline2Block.containsBlockLevelContent(val))
                        {
                            Inline2Block i2b = new Inline2Block();
                            cc = i2b.convertToBlockLevel(cc, true);

                            if (cc == null)
                            {
                                MessageBox.Show("Problems inserting block level XHTML at this location.");
                                return;
                            }

                        }
                    }

                    xppe.setup(null, cxpId, strXPath, prefixMappings, true);
                    xppe.save();

                    td.set("od:xpath", xppe.xpathId);

                    cc.Tag = td.asQueryString();

                } else
                {
                    xppe.setup(null, cxpId, strXPath, prefixMappings, true);
                    xppe.save();

                    td.set("od:xpath", xppe.xpathId);

                    cc.Title = "Data value: " + xppe.xpathId;
                    cc.Tag = td.asQueryString();

                    log.Warn(" XPath \n\r " + strXPath 
                        + "\n\r does not return an element. The OpenDoPE pre-processor will attempt to evaluate it, but Word will never update the result. ");
                    bind.Delete();
                    MessageBox.Show(" XPath \n\r " + strXPath
                        + "\n\r does not return an element. Check this is what you want? ");
                }

            }

        }

        // FIXME - can only get mapping if you know the prefix?!
        static string GetPrefixMappings(Office.CustomXMLPrefixMappings prefixMappings)
        {
            string s = "";
            for (int i = 0; i < prefixMappings.Count; i++)
            {
                Office.CustomXMLPrefixMapping pm = prefixMappings[i];
                s += "xmlns:" + pm.Prefix + "='" + pm.NamespaceURI + "' ";
            }
            return s;
        }


        private void buttonPartSelect_Click(object sender, RibbonControlEventArgs e)
        {
            CustomTaskPane ctpPaneForThisWindow = Utilities.FindTaskPaneForCurrentWindow();
            Controls.ControlMain ccm = (Controls.ControlMain)ctpPaneForThisWindow.Control;
            ccm.formPartList.Show();
        }

        private void toggleButtonDesignMode_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document docx = Globals.ThisAddIn.Application.ActiveDocument;
            //if (!docx.FormsDesign)
            //{
                docx.ToggleFormsDesign();
            //}

        }

        private void buttonXmlOptions_Click(object sender, RibbonControlEventArgs e)
        {
            using (Forms.FormOptions fo = new Forms.FormOptions())
            {
                if (fo.ShowDialog() == DialogResult.OK)
                {
                    ThisAddIn.UpdateSettings(fo.NewOptions);
                }
            }
        }

        private void buttonDelete_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;

            // Get the content control
            Word.ContentControl cc = ContentControlMaker.getActiveContentControl(document, Globals.ThisAddIn.Application.Selection);
            if (cc == null)
            {
                // Shouldn't happen
                MessageBox.Show("Which content control?");
                return;
            }
            // Delete the content control, but keep the contents
            log.Warn("Deleting control with tag " + cc.Tag);
            cc.Delete(false);
        }

        private void buttonAbout_Click(object sender, RibbonControlEventArgs e)
        {
            Forms.FormAbout fo = new Forms.FormAbout();
            fo.ShowDialog();

        }


    }
}
