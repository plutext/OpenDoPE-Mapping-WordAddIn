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
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

using Word = Microsoft.Office.Interop.Word;
using NLog;
using OpenDoPEModel;
using Microsoft.Office.Tools.Word.Extensions; // for VSTOObject, 
using Microsoft.Office.Tools; // for CTP
using System.Windows.Forms;



namespace XmlMappingTaskPane
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {

        static Logger log = LogManager.GetLogger("Ribbon");

        public Ribbon()
        {
            ribbon = this;
        }

        public static Ribbon ribbon;

        public static void myInvalidate()
        {
            //ctp = findCustomTaskPane();
            Ribbon.ribbonUI.Invalidate();
        }

        static private Office.IRibbonUI ribbonUI;
        //public static Office.IRibbonUI getRibbon()
        //{
        //    return ribbon;
        //}

        //public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        //{
        //    ribbon = ribbonUI;
        //}

        public bool buttonBindEnabled = false;
        public bool isButtonBindEnabled(Office.IRibbonControl control) { return buttonBindEnabled; }

        public bool buttonConditionEnabled = false;
        public bool isButtonConditionEnabled(Office.IRibbonControl control) { return buttonConditionEnabled; }

        public bool buttonRepeatEnabled = false;
        public bool isButtonRepeatEnabled(Office.IRibbonControl control) { return buttonRepeatEnabled; }

        public bool buttonEditEnabled = false;
        public bool isButtonEditEnabled(Office.IRibbonControl control) { return buttonEditEnabled; }

        public bool buttonDeleteEnabled = false;
        public bool isButtonDeleteEnabled(Office.IRibbonControl control) { return buttonDeleteEnabled; }

        public bool menuAdvancedEnabled = false;

        public bool buttonReplaceXMLEnabled = false;
        public bool isButtonReplaceXMLEnabled(Office.IRibbonControl control) { return buttonReplaceXMLEnabled;}

        public bool buttonClearAllEnabled = false;
        public bool isButtonClearAllEnabled(Office.IRibbonControl control) { return buttonClearAllEnabled; }

        public bool toggleButtonMappingChecked = false;
        public bool toggleButtonMappingEnabled = false; // TODO REVIEW

        public bool isContextMenuStylesOverridden(Office.IRibbonControl control) {
            if (!buttonEditEnabled)
            {
                log.Debug("edit disabled");
                return false; // outside any cc
            }
            // we should be in a content control
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            Word.ContentControl cc = ContentControlMaker.getActiveContentControl(document, Globals.ThisAddIn.Application.Selection);
            if (cc == null)
            {
                log.Debug("cc null");
                return false;
            }
            // we are in a cc. is it an XHTML one?
            log.Debug(cc.Title);
            log.Debug(cc.Tag);
            return ((cc.Title != null && cc.Title.StartsWith("XHTML"))
                || (cc.Tag != null && cc.Tag.Contains("ContentType=application/xhtml+xml")));
        }

        public bool isContextMenuStyleRetained(Office.IRibbonControl control) { 
            return !isContextMenuStylesOverridden(control);
        }


        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("XmlMappingTaskPane.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            //this.ribbon = ribbonUI;
            Ribbon.ribbonUI = ribbonUI;
        }

        public void ContextMenuText_odStyles_Click(Office.IRibbonControl control)
        {
            log.Debug("Clicked!");
            //MessageBox.Show("Clicked!");

            Forms.FormStyleChooser fsc = new Forms.FormStyleChooser();
            DialogResult result = fsc.ShowDialog();
            if (result.Equals(DialogResult.Cancel))
            {
                // restore previous value
                fsc.revert();
            }
        }


        public void toggleButtonMapping_Click(Office.IRibbonControl control, bool pressed)//, ref bool cancelDefault)
        {
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;

            //get the ctp for this window (or null if there's not one)
            CustomTaskPane ctpPaneForThisWindow = Utilities.FindTaskPaneForCurrentWindow();

            if (pressed) toggleButtonMappingChecked = true;


            if (toggleButtonMappingChecked == false)
            {
                //Debug.Assert(ctpPaneForThisWindow != null,
                //    "how was the ribbon button pressed if there was a control?");

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
                        toggleButtonMappingChecked = false;
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
        public void buttonReplaceXML_Click(Office.IRibbonControl control)
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

        public void buttonClearAll_Click(Office.IRibbonControl control)
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
                foreach (Word.HeaderFooter hf in section.Headers)
                {
                    foreach (Word.ContentControl cc in hf.Range.ContentControls)
                    {
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
            this.buttonBindEnabled = false;
            this.buttonConditionEnabled = false;
            this.buttonRepeatEnabled = false;

            this.buttonEditEnabled = false;
            this.buttonDeleteEnabled = false;

            this.menuAdvancedEnabled = false;

            this.buttonReplaceXMLEnabled = false;
            this.buttonClearAllEnabled = false;

            toggleButtonMappingChecked = false;


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
            de.Ribbon = this; // so de can enable/disable stuff on ribbon


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

            this.buttonBindEnabled = true;
            this.buttonConditionEnabled = true;
            this.buttonRepeatEnabled = true;

            this.menuAdvancedEnabled = true;

            this.buttonReplaceXMLEnabled = true;
            this.buttonClearAllEnabled = true;

            myInvalidate();

            //reset the cursor
            Globals.ThisAddIn.Application.System.Cursor = Word.WdCursorType.wdCursorNormal;

        }

        public void buttonBind_Click(Office.IRibbonControl control)
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

        public void buttonCondition_Click(Office.IRibbonControl control)
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
                    document.Windows[1].Selection.Text = "condition";

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
        public void buttonRepeat_Click(Office.IRibbonControl control)
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
            catch (System.Exception)
            {
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
            if (thisNodeSibling != null
                && thisNodeSibling.BaseName.Equals(thisNode.BaseName))
            {
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
            if (strXPath.EndsWith("]"))
            {
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


        public void buttonEdit_Click(Office.IRibbonControl control)
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
            object missing = System.Type.Missing;

            // First, work out whether this is a condition or a repeat or a plain bind
            bool isCondition = false;
            bool isRepeat = false;
            bool isBind = false;
            if ((cc.Title != null && cc.Title.StartsWith("Condition"))
                || (cc.Tag != null && cc.Tag.Contains("od:condition")))
            {
                isCondition = true;
            }
            else if ((cc.Title != null && cc.Title.StartsWith("Repeat"))
                || (cc.Tag != null && cc.Tag.Contains("od:repeat")))
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
                if (c != null
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
            else if (isBind)
            {

                if (cc.XMLMapping.IsMapped)
                {
                    // Prefer this, if for some reason it contradicts od:xpath
                    strXPath = cc.XMLMapping.XPath;
                    cxpId = cc.XMLMapping.CustomXMLPart.Id;
                    prefixMappings = cc.XMLMapping.PrefixMappings;

                }
                else if (td.get("od:xpath") != null)
                {
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
                bool mappable = false;
                try
                {
                    mappable = bind.SetMapping(strXPath, prefixMappings,
                        CustomXmlUtilities.getPartById(document, cxpId));
                }
                catch (COMException ce)
                {
                    if (ce.Message.Contains("Data bindings cannot be created for rich text content controls"))
                    {
                        // TODO: editing a rich text control
                        // TODO manually check whether it is mappable
                        // So for now, 
                        cc.Delete(true);
                        cc = document.ContentControls.Add(
                            Word.WdContentControlType.wdContentControlText, ref missing);
                        mappable = cc.XMLMapping.SetMapping(strXPath, prefixMappings,
                            CustomXmlUtilities.getPartById(document, cxpId));
                    }
                    else
                    {
                        log.Error(ce);
                        //What to do??
                    }
                }
                if (mappable)
                {
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
                        Globals.ThisAddIn.Application.Selection.Collapse(ref missing);

                        bool _PictureContentControlsReplace = true;
                        String picSetting = System.Configuration.ConfigurationManager.AppSettings["ContentControl.Picture.RichText.Override"];
                        if (picSetting != null)
                        {
                            Boolean.TryParse(picSetting, out _PictureContentControlsReplace);
                        }

                        if (_PictureContentControlsReplace)
                        {
                            // Use a rich text control instead
                            cc = document.ContentControls.Add(
                                Word.WdContentControlType.wdContentControlRichText, ref missing);

                            PictureUtils.setPictureHandler(td);
                            cc.Title = "Image: " + xppe.xpathId;

                            PictureUtils.pastePictureIntoCC(cc, Convert.FromBase64String(val));
                        }
                        else
                        {
                            cc = document.ContentControls.Add(
                                Word.WdContentControlType.wdContentControlPicture, ref missing);

                            cc.XMLMapping.SetMapping(strXPath, prefixMappings,
                                CustomXmlUtilities.getPartById(document, cxpId));
                        }

                    }
                    else if (ContentDetection.IsFlatOPCContent(val))
                    {
                        // <?mso-application progid="Word.Document"?>
                        // <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">

                        td.set("od:progid", "Word.Document");
                        cc.Tag = td.asQueryString();

                        cc.XMLMapping.Delete();
                        cc.Type = Word.WdContentControlType.wdContentControlRichText;
                        cc.Title = "Word: " + xppe.xpathId;
                        //cc.Range.Text = val; // don't escape it

                        Inline2Block i2b = new Inline2Block();
                        cc = i2b.convertToBlockLevel(cc, false, true);

                        if (cc == null)
                        {
                            MessageBox.Show("Problems inserting block level WordML at this location.");
                            return;
                        }

                        cc.Range.InsertXML(val, ref missing);
                    } 
                    else if (ContentDetection.IsXHTMLContent(val))
                    {
                        td.set("od:ContentType", "application/xhtml+xml");
                        cc.Tag = td.asQueryString();

                        cc.XMLMapping.Delete();
                        cc.Type = Word.WdContentControlType.wdContentControlRichText;
                        cc.Title = "XHTML: " + xppe.xpathId;

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

                    xppe.setup(null, cxpId, strXPath, prefixMappings, true);
                    xppe.save();

                    td.set("od:xpath", xppe.xpathId);

                    cc.Tag = td.asQueryString();

                }
                else
                {
                    xppe.setup(null, cxpId, strXPath, prefixMappings, true);
                    xppe.save();

                    td.set("od:xpath", xppe.xpathId);

                    cc.Title = "Data value: " + xppe.xpathId;
                    cc.Tag = td.asQueryString();

                    // FIXME TODO handle pictures/FlatOPC/XHTML in this case

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


        public void buttonPartSelect_Click(Office.IRibbonControl control)
        {
            CustomTaskPane ctpPaneForThisWindow = Utilities.FindTaskPaneForCurrentWindow();
            Controls.ControlMain ccm = (Controls.ControlMain)ctpPaneForThisWindow.Control;
            ccm.formPartList.Show();
        }

        public void toggleButtonDesignMode_Click(Office.IRibbonControl control, bool pressed)
        {
            Word.Document docx = Globals.ThisAddIn.Application.ActiveDocument;
            //if (!docx.FormsDesign)
            //{
            docx.ToggleFormsDesign();
            //}

        }

        public void buttonXmlOptions_Click(Office.IRibbonControl control)
        {
            using (Forms.FormOptions fo = new Forms.FormOptions())
            {
                if (fo.ShowDialog() == DialogResult.OK)
                {
                    ThisAddIn.UpdateSettings(fo.NewOptions);
                }
            }
        }

        public void buttonDelete_Click(Office.IRibbonControl control)
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

        public void buttonAbout_Click(Office.IRibbonControl control)
        {
            Forms.FormAbout fo = new Forms.FormAbout();
            fo.ShowDialog();

        }


        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
