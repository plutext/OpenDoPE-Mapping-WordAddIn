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
using System;
using System.Collections;
using System.Diagnostics;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Word.Extensions; // for VSTOObject
using NLog;
using OpenDoPEModel;
using System.Threading;

namespace XmlMappingTaskPane
{
    class DocumentEvents
    {
        static Logger log = LogManager.GetLogger("DocumentEvents");

        //task pane control
        private Controls.ControlMain m_cmTaskPane;

        //current part
        private Word.Document m_wddoc;
        private Office.CustomXMLParts m_parts;
        private Office.CustomXMLPart m_currentPart;

        //document/streams/stream event handler storage varaibles
        private ArrayList m_alDocumentEvents = new ArrayList();
        private ArrayList m_alPartsEvents = new ArrayList();
        private ArrayList m_alPartEvents = new ArrayList();

        public DocumentEvents(Controls.ControlMain cm)
        {
            //get the initial part
            m_cmTaskPane = cm;
            m_wddoc = Globals.ThisAddIn.Application.ActiveDocument;
            m_parts = m_wddoc.CustomXMLParts;
            m_currentPart = cm.model.getUserPart(System.Configuration.ConfigurationManager.AppSettings["RootElement"]);

            //hook up event handlers
            SetupEventHandlers();
        }

        /// <summary>
        /// A reference to the RibbonMapping object, so
        /// we can enable/disable content control related
        /// buttons depending on whether we are in a content 
        /// control or not (since the content control enter/
        /// exit events are in this class).
        /// </summary>
        public Ribbon Ribbon { get; set; }

        /// <summary>
        /// Set up the document-level event handlers.
        /// </summary>
        private void SetupEventHandlers()
        {

            //add the new document level event handlers
            m_alDocumentEvents.Add(
                new Word.DocumentEvents2_ContentControlOnEnterEventHandler(
                    doc_ContentControlOnEnter));
            m_alDocumentEvents.Add(
                new Word.DocumentEvents2_ContentControlAfterAddEventHandler(
                    doc_ContentControlAfterAdd));
            m_alDocumentEvents.Add(new Word.DocumentEvents2_ContentControlOnExitEventHandler(doc_ContentControlOnExit));

            m_wddoc.ContentControlOnEnter
                += (Word.DocumentEvents2_ContentControlOnEnterEventHandler)m_alDocumentEvents[0];
            m_wddoc.ContentControlAfterAdd
                += (Word.DocumentEvents2_ContentControlAfterAddEventHandler)m_alDocumentEvents[1];
            m_wddoc.ContentControlOnExit += (Word.DocumentEvents2_ContentControlOnExitEventHandler)m_alDocumentEvents[2];

            //set up stream event handlers for this document
            m_alPartsEvents.Add(new Office._CustomXMLPartsEvents_PartAfterAddEventHandler(parts_PartAdd));
            m_alPartsEvents.Add(new Office._CustomXMLPartsEvents_PartAfterLoadEventHandler(parts_PartLoad));
            m_alPartsEvents.Add(new Office._CustomXMLPartsEvents_PartBeforeDeleteEventHandler(parts_PartDelete));
            m_parts.PartAfterAdd += (Office._CustomXMLPartsEvents_PartAfterAddEventHandler)m_alPartsEvents[0];
            m_parts.PartAfterLoad += (Office._CustomXMLPartsEvents_PartAfterLoadEventHandler)m_alPartsEvents[1];
            m_parts.PartBeforeDelete += (Office._CustomXMLPartsEvents_PartBeforeDeleteEventHandler)m_alPartsEvents[2];

            //set up event handlers for the first stream (shown by default)
            m_alPartEvents.Add(new Office._CustomXMLPartEvents_NodeAfterDeleteEventHandler(part_NodeAfterDelete));
            m_alPartEvents.Add(new Office._CustomXMLPartEvents_NodeAfterInsertEventHandler(part_NodeAfterInsert));
            m_alPartEvents.Add(new Office._CustomXMLPartEvents_NodeAfterReplaceEventHandler(part_NodeAfterReplace));
            m_currentPart.NodeAfterDelete += (Office._CustomXMLPartEvents_NodeAfterDeleteEventHandler)m_alPartEvents[0];
            m_currentPart.NodeAfterInsert += (Office._CustomXMLPartEvents_NodeAfterInsertEventHandler)m_alPartEvents[1];
            m_currentPart.NodeAfterReplace += (Office._CustomXMLPartEvents_NodeAfterReplaceEventHandler)m_alPartEvents[2];
        }


        /// <summary>
        /// Change the currently active document.
        /// </summary>
        internal void ChangeCurrentDocument()
        {
            log.Debug("Changing event document to " + Globals.ThisAddIn.Application.ActiveDocument.Name);

            DisconnectOldDocument();
            ConnectCurrent();
        }

        public void NodeAfterReplaceDisconnect()
        {
            m_currentPart.NodeAfterReplace -= (Office._CustomXMLPartEvents_NodeAfterReplaceEventHandler)m_alPartEvents[2];
        }

        public void NodeAfterReplaceReconnect()
        {
            m_currentPart.NodeAfterReplace += (Office._CustomXMLPartEvents_NodeAfterReplaceEventHandler)m_alPartEvents[2];
        }

        public void DisconnectOldDocument() {

            //unhook existing events
            m_wddoc.ContentControlOnEnter -= (Word.DocumentEvents2_ContentControlOnEnterEventHandler)m_alDocumentEvents[0];
            m_wddoc.ContentControlAfterAdd -= (Word.DocumentEvents2_ContentControlAfterAddEventHandler)m_alDocumentEvents[1];
            m_wddoc.ContentControlOnExit -= (Word.DocumentEvents2_ContentControlOnExitEventHandler)m_alDocumentEvents[2];


            m_parts.PartAfterAdd -= (Office._CustomXMLPartsEvents_PartAfterAddEventHandler)m_alPartsEvents[0];
            m_parts.PartAfterLoad -= (Office._CustomXMLPartsEvents_PartAfterLoadEventHandler)m_alPartsEvents[1];
            m_parts.PartBeforeDelete -= (Office._CustomXMLPartsEvents_PartBeforeDeleteEventHandler)m_alPartsEvents[2];
            m_currentPart.NodeAfterDelete -= (Office._CustomXMLPartEvents_NodeAfterDeleteEventHandler)m_alPartEvents[0];
            m_currentPart.NodeAfterInsert -= (Office._CustomXMLPartEvents_NodeAfterInsertEventHandler)m_alPartEvents[1];
            m_currentPart.NodeAfterReplace -= (Office._CustomXMLPartEvents_NodeAfterReplaceEventHandler)m_alPartEvents[2];

            //release the streams + stream handler references
            m_alDocumentEvents.Clear();
            m_alPartsEvents.Clear();
            m_alPartEvents.Clear();

            //clean up the m_wddoc object (since otherwise the RCW gets disposed out from under it)
            m_wddoc = null;
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

        }

        internal void ConnectCurrent() {

            //hook up new objects
            m_wddoc = Globals.ThisAddIn.Application.ActiveDocument;
            m_parts = m_wddoc.CustomXMLParts;
            //m_currentPart = m_parts[1];  // bad that this is done here as well!

            m_cmTaskPane.model = OpenDoPEModel.Model.ModelFactory(m_wddoc);
            
            m_cmTaskPane.formPartList.Dispose();
            m_cmTaskPane.formPartList = new Forms.FormSwitchSelectedPart();
            m_cmTaskPane.formPartList.controlPartList.controlMain = m_cmTaskPane;

            if (m_cmTaskPane.model.userParts.Count == 0)
            {
                log.Error("No users part found! This shouldn't happen!");
            }
            m_currentPart = m_cmTaskPane.model.userParts[0];


            SetupEventHandlers();
        }

        /// <summary>
        /// Change the currently active XML part.
        /// </summary>
        /// <param name="cxp">The CustomXMLPart specifying the newly active XML part.</param>
        internal void ChangeCurrentPart(Office.CustomXMLPart cxp)
        {
            if (cxp != null)
            {
                log.Debug("Changing event stream to " + cxp.DocumentElement.BaseName + ".."  +   cxp.Id);

                //unhook the stream event handlers
                m_currentPart.NodeAfterDelete -= (Office._CustomXMLPartEvents_NodeAfterDeleteEventHandler)m_alPartEvents[0];
                m_currentPart.NodeAfterInsert -= (Office._CustomXMLPartEvents_NodeAfterInsertEventHandler)m_alPartEvents[1];
                m_currentPart.NodeAfterReplace -= (Office._CustomXMLPartEvents_NodeAfterReplaceEventHandler)m_alPartEvents[2];

                //release the streams + event handler references
                m_currentPart = null;
                m_alPartEvents.Clear();

                //hook up the new stream
                m_currentPart = cxp;

                //set up event handlers on the supplied stream
                m_alPartEvents.Add(new Office._CustomXMLPartEvents_NodeAfterDeleteEventHandler(part_NodeAfterDelete));
                m_alPartEvents.Add(new Office._CustomXMLPartEvents_NodeAfterInsertEventHandler(part_NodeAfterInsert));
                m_alPartEvents.Add(new Office._CustomXMLPartEvents_NodeAfterReplaceEventHandler(part_NodeAfterReplace));
                m_currentPart.NodeAfterDelete += (Office._CustomXMLPartEvents_NodeAfterDeleteEventHandler)m_alPartEvents[0];
                m_currentPart.NodeAfterInsert += (Office._CustomXMLPartEvents_NodeAfterInsertEventHandler)m_alPartEvents[1];
                m_currentPart.NodeAfterReplace += (Office._CustomXMLPartEvents_NodeAfterReplaceEventHandler)m_alPartEvents[2];
            }
            else
            {
                log.Error("SetCurrentStream received a null stream");
            }
        }

        #region Document-level events

        /// <summary>
        /// Handle Word's OnEnter event for content controls, to set the selection in the pane (if the option is set).
        /// </summary>
        /// <param name="ccEntered">A ContentControl object specifying the control that was entered.</param>
        private void doc_ContentControlOnEnter(Word.ContentControl ccEntered)
        {
            log.Debug("Document.ContentControlOnEnter fired.");
            if (ccEntered.XMLMapping.IsMapped)
            {
                log.Debug("control mapped to " + ccEntered.XMLMapping.XPath);
                if (ccEntered.XMLMapping.CustomXMLNode != null)
                {
                    m_cmTaskPane.RefreshControls(Controls.ControlMain.ChangeReason.OnEnter, null, null, null, ccEntered.XMLMapping.CustomXMLNode, null);
                }
                else
                {
                    // Not mapped to anything; probably because the part was replaced?
                    log.Debug(".. but XMLMapping.CustomXMLNode is null");
                    m_cmTaskPane.controlTreeView.DeselectNode();
                    m_cmTaskPane.WarnViaProperties(ccEntered.XMLMapping.XPath);
                }
            }
            else if (ccEntered.Tag!=null)
            {
                
                TagData td = new TagData(ccEntered.Tag);
                string xpathid = td.getXPathID();
                if (xpathid==null) {
                    xpathid = td.getRepeatID();
                }

                if (xpathid == null)
                {
                    // Visually indicate in the task pane that we're no longer in a mapped control
                    m_cmTaskPane.controlTreeView.DeselectNode();
                    m_cmTaskPane.PropertiesClear();

                } else {
                    log.Debug("control mapped via tag to " + xpathid);

                    // Repeats, escaped XHTML; we don't show anything for conditions
                    XPathsPartEntry xppe = new XPathsPartEntry(m_cmTaskPane.model);
                    xpathsXpath xx = xppe.getXPathByID(xpathid);
                    Office.CustomXMLNode customXMLNode = null;
                    if (xx != null)
                    {
                        log.Debug(xx.dataBinding.xpath);
                        customXMLNode = m_currentPart.SelectSingleNode(xx.dataBinding.xpath);
                    }
                    if (customXMLNode == null)
                    {
                        // Not mapped to anything; probably because the part was replaced?
                        m_cmTaskPane.controlTreeView.DeselectNode();
                        if (xx == null)
                        {
                            log.Error("Couldn't find xpath for " + xpathid);
                            m_cmTaskPane.WarnViaProperties("Missing");
                        }
                        else
                        {
                            log.Warn("Couldn't find target node for " + xx.dataBinding.xpath);
                            m_cmTaskPane.WarnViaProperties(xx.dataBinding.xpath);
                        }
                    }
                    else
                    {
                        m_cmTaskPane.RefreshControls(Controls.ControlMain.ChangeReason.OnEnter, null, null, null, customXMLNode, null);
                    }
                }
            }
            Ribbon.buttonEditEnabled = true;
            Ribbon.buttonDeleteEnabled = true;

            Ribbon.myInvalidate();
        }

        private void doc_ContentControlOnExit(Word.ContentControl ContentControl, ref bool Cancel) 
        {
            log.Debug("doc_ContentControlOnExit fired.");
            //Cancel = false;

            Ribbon.buttonEditEnabled = false;
            Ribbon.buttonDeleteEnabled = false;

            // Visually indicate in the task pane that we're no longer in a mapped control
            m_cmTaskPane.controlTreeView.DeselectNode();
            m_cmTaskPane.PropertiesClear();

            // Note that this event does not fire when a selected control
            // (created by drag) loses focus after user clicks elsewhere.
            // So more would be required to clear the properties in this case.


            // TODO: only do that if not in another content control?
            // Doesn't seem to be necessary


            // Necessary only in the case where the clipboard is used
            //ContentControlStyle.CopyAdjacentFormat(ContentControl);

            Ribbon.myInvalidate();

        }

        /// <summary>
        /// Handle Word's AfterAdd event for content controls, to set new controls' placeholder text
        /// </summary>
        /// <param name="ccAdded"></param>
        /// <param name="InUndoRedo"></param>
        private void doc_ContentControlAfterAdd(Word.ContentControl ccAdded, bool InUndoRedo)
        {
            log.Info("doc_ContentControlAfterAdd...");

            if (InUndoRedo) {
            }  else if (m_cmTaskPane.RecentDragDrop)
            {

                log.Info("recent drag drop...");
                 
                ccAdded.Application.ScreenUpdating = false;

                /*
                 * Don't set placeholder text.  We don't want it!
                 * It prevents child content controls from being added.
                 * And it is very difficult to get rid of!
                 * 
                //set the placeholder text 
                if (m_cmTaskPane.controlMode1.isModeBind()
                    && ccAdded.Type != Word.WdContentControlType.wdContentControlRichText
                    && ccAdded.Type != Word.WdContentControlType.wdContentControlPicture  // TODO, what other types don't support placeholder text?
                    )
                {
                    log.Debug("Setting placeholder text (fwiw)");
                    // set the placeholder text has the side effect of clearing out the control's contents,
                    // so grab the current text in the node (if any)
                    string currentText = null;
                    if (ccAdded.XMLMapping.IsMapped)
                    {
                        currentText = ccAdded.XMLMapping.CustomXMLNode.Text;
                    }

                    ccAdded.SetPlaceholderText(null, null, Utilities.GetPlaceholderText(ccAdded.Type));

                    // now bring back the original text
                    if (currentText != null)
                    {
                        ccAdded.Range.Text = currentText;
                    }

                }
                 * */


                string xml = ccAdded.Range.Text;
                if (ContentDetection.IsBase64Encoded(xml))
                {
                    // Don't need to do anything here...

                } else  if (ContentDetection.IsFlatOPCContent(xml))
                {
                    xml = xml.Replace("&quot;", "\"");
                    log.Debug(xml);


                    // Need it to be block level
                    Inline2Block i2b = new Inline2Block();
                    AsyncMethodCallerFlatOPC caller = new AsyncMethodCallerFlatOPC(i2b.blockLevelFlatOPC);
                    caller.BeginInvoke(ccAdded, xml, null, null);

                } else if (ContentDetection.IsXHTMLContent(ccAdded.Range.Text))
                {

                    log.Info("is XHTML .. " + ccAdded.Range.Text);

                    if (Inline2Block.containsBlockLevelContent(ccAdded.Range.Text))
                    {
                        Inline2Block i2b = new Inline2Block();
                        AsyncMethodCallerXHTML caller = new AsyncMethodCallerXHTML(i2b.convertToBlockLevel);
                        caller.BeginInvoke(ccAdded, true, true, null, null);
                    }

                }
                else
                {
                    // Have to do CopyAdjacentFormat outside this event.
                    // (It works from OnExit, but not from AfterAdd. Go figure...)
                    ContentControlStyle ccs = new ContentControlStyle();
                    AsyncMethodCallerPlainText caller2 = new AsyncMethodCallerPlainText(ccs.CopyAdjacentFormat);
                    caller2.BeginInvoke(ccAdded, null, null);
                }
            }
        }



        public delegate Word.ContentControl AsyncMethodCallerXHTML(Word.ContentControl cc, bool foo, bool updateScreen);
        public delegate Word.ContentControl AsyncMethodCallerFlatOPC(Word.ContentControl cc, string xml);
        public delegate Word.ContentControl AsyncMethodCallerPlainText(Word.ContentControl cc);

        #endregion

        #region XML part-level events

        private void parts_PartAdd(Office.CustomXMLPart NewStream)
        {
            log.Debug("Streams.StreamAfterAdd fired.");
            m_cmTaskPane.RefreshControls(Controls.ControlMain.ChangeReason.PartAdded, null, null, null, null, null);
        }

        /// <summary>
        /// This can't happen, since we don't provide the user with a way to delete a part,
        /// othen than RibbonMapping buttonClearAll_Click 
        /// </summary>
        /// <param name="OldStream"></param>
        private void parts_PartDelete(Office.CustomXMLPart OldStream)
        {
            log.Debug("Streams.StreamBeforeDelete fired.");
            m_cmTaskPane.RefreshControls(Controls.ControlMain.ChangeReason.PartDeleted, null, null, null, null, OldStream);
        }

        private void parts_PartLoad(Office.CustomXMLPart LoadedStream)
        {
            log.Debug("Streams.StreamAfterLoad fired.");
            m_cmTaskPane.RefreshControls(Controls.ControlMain.ChangeReason.PartLoaded, null, null, null, null, null);
        }

        private void part_NodeAfterDelete(Office.CustomXMLNode mxnDeletedNode, Office.CustomXMLNode mxnDeletedParent, Office.CustomXMLNode mxnDeletedNextSibling, bool bInUndoRedo)
        {
            log.Debug("Streams.NodeAfterDelete fired.");
            m_cmTaskPane.RefreshControls(Controls.ControlMain.ChangeReason.NodeDeleted, mxnDeletedNode, mxnDeletedParent, mxnDeletedNextSibling, null, null);
        }

        private void part_NodeAfterInsert(Office.CustomXMLNode mxnNewNode, bool bInUndoRedo)
        {
            log.Debug("Streams.NodeAfterInsert fired.");
            m_cmTaskPane.RefreshControls(Controls.ControlMain.ChangeReason.NodeAdded, null, null, null, mxnNewNode, null);
        }

        private void part_NodeAfterReplace(Office.CustomXMLNode mxnOldNode, Office.CustomXMLNode mxnNewNode, bool bInUndoRedo)
        {
            log.Debug("Streams.NodeAfterReplace fired.");
            m_cmTaskPane.RefreshControls(Controls.ControlMain.ChangeReason.NodeReplaced, mxnOldNode, mxnNewNode.ParentNode, mxnNewNode.NextSibling, mxnNewNode, null);
        }

        #endregion

        /// <summary>
        /// Get the currently active XML part. Read-only.
        /// </summary>
        internal Office.CustomXMLPart Part
        {
            get
            {
                return m_currentPart;
            }
        }

        /// <summary>
        /// Get the XML part collection for the current document. Read-only.
        /// </summary>
        internal Office.CustomXMLParts PartCollection
        {
            get
            {
                return m_parts;
            }
        }

        /// <summary>
        /// Get the currently active document. Read-only.
        /// </summary>
        internal Word.Document Document
        {
            get
            {
                return m_wddoc;
            }
        }
    }
}
