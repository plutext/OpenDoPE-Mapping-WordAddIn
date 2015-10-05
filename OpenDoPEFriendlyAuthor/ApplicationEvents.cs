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
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using NLog;

namespace XmlMappingTaskPane
{
    /// <summary>
    /// DocumentChange, WindowActivate
    /// </summary>
    class ApplicationEvents
    {
        static Logger log = LogManager.GetLogger("ApplicationEvents");

        /// <summary>
        /// true (default) if there is a document frame window
        /// for each open document
        /// </summary>
        private bool m_fLastSdiMode;

        private int m_intLastDocumentCount; 
        private int m_intLastWindowCount;

        //private RibbonToggleButton m_buttonMapping;
        private Ribbon m_ribbon;
        private IDictionary<Word.Window, CustomTaskPane> m_dictTaskPanes;

        private Word.ApplicationEvents4_DocumentChangeEventHandler m_ehDocumentChange;
        private Word.ApplicationEvents4_WindowActivateEventHandler m_ehWindowActivate;
        private Word.Window m_wdwinLastWindow;

        internal ApplicationEvents(Ribbon rm, IDictionary<Word.Window, CustomTaskPane> dictTaskPanes)
        {
            //store the application and Ribbon objects
            //m_buttonMapping = rm.toggleButtonMapping;
            m_ribbon = rm;
            m_dictTaskPanes = dictTaskPanes;

            //store MDI/SDI state
            m_fLastSdiMode = Globals.ThisAddIn.Application.ShowWindowsInTaskbar;
            log.Debug("ShowWindowsInTaskbar? " + m_fLastSdiMode);

            //capture the necessary app events
            m_ehDocumentChange = new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentChangeEventHandler(app_DocumentChange);
            m_ehWindowActivate = new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowActivateEventHandler(app_WindowActivate);
            Globals.ThisAddIn.Application.DocumentChange += m_ehDocumentChange;
            Globals.ThisAddIn.Application.WindowActivate += m_ehWindowActivate;
        }

        /// <summary>
        /// Handle Word's DocumentChange event (a new document has focus).
        /// </summary>
        private void app_DocumentChange()
        {
            log.Debug("app_DocumentChange() fired");

            //refresh the ribbon button
            UpdateButtonState();

            int intNewDocCount = Globals.ThisAddIn.Application.Documents.Count;

            if (intNewDocCount == 0 && m_intLastDocumentCount == 1)
            {
                //check if we hit the fishbowl
                //delete the current CTP (it should always be the last one)
                Debug.Assert(Globals.ThisAddIn.CustomTaskPanes.Count <= 1, "why are there undeleted CTPs around?");
                if (Globals.ThisAddIn.CustomTaskPanes.Count == 1)
                {
                    Globals.ThisAddIn.CustomTaskPanes.RemoveAt(0);
                }  
            }

            //in MDI, need to update the task pane to show the content for the new document
            if (!Globals.ThisAddIn.Application.ShowWindowsInTaskbar 
                && Globals.ThisAddIn.Application.Documents.Count > 0 
                && Globals.ThisAddIn.CustomTaskPanes.Count > 0) 
                //we're in MDI and there are documents and task panes created
            {
                ((Controls.ControlMain)Globals.ThisAddIn.CustomTaskPanes[0].Control).RefreshControls(Controls.ControlMain.ChangeReason.DocumentChanged, null, null, null, null, null);
                // TODO: review whether this is enough
            }

            //save the new document count
            m_intLastDocumentCount = intNewDocCount;
        }

        /// <summary>
        /// Handle Word's WindowActivate event (a new window has focus).
        /// </summary>
        /// <param name="wddoc">A Document object specifying the document that has focus.</param>
        /// <param name="wdwin">A Window object specifying the window that has focus.</param>
        private void app_WindowActivate(Word.Document wddoc, Word.Window wdwin)
        {
            log.Debug("app_WindowActivate fired");

            int intNewWindowCount = Globals.ThisAddIn.Application.Windows.Count;

            //check if we lost a window, if so, clean up its CustomTaskPane
            if (intNewWindowCount - m_intLastWindowCount == -1)
            {
                //a window was closed - remove any lingering CTP from the collection
                if (m_dictTaskPanes.ContainsKey(m_wdwinLastWindow))
                {
                    Globals.ThisAddIn.CustomTaskPanes.Remove(m_dictTaskPanes[m_wdwinLastWindow]);
                    m_dictTaskPanes.Remove(m_wdwinLastWindow);
                }
            }

            //check for SDI<-->MDI changes
            CheckForMdiSdiSwitch();

            //store new window and count
            m_intLastWindowCount = intNewWindowCount;
            m_wdwinLastWindow = wdwin;

            // If the user did View > New Window,
            // they get a window without a task pane.
            // So disable the buttons
            // (especially the main button), since
            // it is otherwise in a state which
            // contradicts the assertion in RibbonMapping.
            UpdateButtonState();

        }

        private void DisableSecondaryButtons()
        {

            m_ribbon.buttonBindEnabled = false;
            m_ribbon.buttonConditionEnabled = false;
            m_ribbon.buttonRepeatEnabled = false;

            //m_ribbon.buttonEdit.Enabled = false;
            //m_ribbon.buttonDelete.Enabled = false;

            m_ribbon.menuAdvancedEnabled = false;
        }

        private void EnableSecondaryButtons()
        {

            m_ribbon.buttonBindEnabled = true;
            m_ribbon.buttonConditionEnabled = true;
            m_ribbon.buttonRepeatEnabled = true;

            m_ribbon.menuAdvancedEnabled = true;
        }

        /// <summary>
        /// Update the state of our button on the Ribbon.
        /// </summary>
        private void UpdateButtonState()
        {
            //check for an MDI<-->SDI switch
            CheckForMdiSdiSwitch();

            try
            {
                //only leave it off in the fishbowl
                if (Globals.ThisAddIn.Application.Documents.Count == 0)
                {
                    m_ribbon.toggleButtonMappingChecked = false;
                    m_ribbon.toggleButtonMappingEnabled = false;

                    DisableSecondaryButtons();

                    Ribbon.myInvalidate();

                    return;
                }
                else
                    m_ribbon.toggleButtonMappingEnabled = true;
            }
            catch (COMException ex)
            {
                Debug.Fail(ex.Source, ex.Message);
                m_ribbon.toggleButtonMappingEnabled = false;
                return;
            }

            //check pressed state
            if (m_fLastSdiMode)
            {
                //get the ctp for this window (or null if there's not one)
                CustomTaskPane ctpPaneForThisWindow = null;
                try
                {
                    Globals.ThisAddIn.TaskPaneList.TryGetValue(Globals.ThisAddIn.Application.ActiveWindow, out ctpPaneForThisWindow);
                }
                catch (COMException ex)
                {
                    if (ex.Message.Contains("no document is open"))
                    {
                        // silently swallow
                        log.Debug("Silently swallowing " + ex.Message);
                        /* To reproduce:
                         * - in a doc1, add content controls
                         * - create a new doc2 (just a plain Word doc)
                         * - close doc1
                         */ 
                        return;
                    }
                    else
                    {
                        Debug.Fail("Failed to get CTP:" + ex.Message);
                    }

                }

                //if it's not built, don't check
                if (ctpPaneForThisWindow == null)
                {
                    m_ribbon.toggleButtonMappingChecked = false;
                    DisableSecondaryButtons();
                }
                else
                {
                    //if it's visible, down
                    if (ctpPaneForThisWindow.Visible == true)
                    {
                        m_ribbon.toggleButtonMappingChecked = true;
                        EnableSecondaryButtons();
                    }
                    else
                    {
                        m_ribbon.toggleButtonMappingChecked = false;
                        DisableSecondaryButtons();
                    }
                }

            }
            else
            {
                if (Globals.ThisAddIn.CustomTaskPanes.Count > 0)
                {
                    Debug.Assert(Globals.ThisAddIn.CustomTaskPanes.Count == 1, "why are there multiple CTPs?");

                    if (Globals.ThisAddIn.CustomTaskPanes[0].Visible)
                    {
                        m_ribbon.toggleButtonMappingChecked = true;
                        EnableSecondaryButtons();
                    }
                    else
                    {
                        m_ribbon.toggleButtonMappingChecked = false;
                        DisableSecondaryButtons();
                    }
                }
            }
            Ribbon.myInvalidate();

        }

        /// <summary>
        /// Check if the application has moved from MDI mode to SDI mode (or vice versa).
        /// </summary>
        private void CheckForMdiSdiSwitch()
        {
            //check if we changed
            log.Debug("ShowWindowsInTaskbar? " + Globals.ThisAddIn.Application.ShowWindowsInTaskbar);

            if (Globals.ThisAddIn.Application.ShowWindowsInTaskbar != m_fLastSdiMode)
            {
                log.Debug(".. which is a change");
                if (m_fLastSdiMode)
                {
                    //going to MDI
                    //the CTP associated with the active window is the only one we need to keep
                    for (int i = Globals.ThisAddIn.CustomTaskPanes.Count - 1; i >= 0; i--)
                    {
                        try
                        {
                            if (Globals.ThisAddIn.CustomTaskPanes[i].Window != Globals.ThisAddIn.Application.ActiveWindow)
                                Globals.ThisAddIn.CustomTaskPanes.RemoveAt(i);
                        }
                        catch (COMException)
                        {
                            //the task pane was disposed by Office, so just remove it from our collection
                            Globals.ThisAddIn.CustomTaskPanes.RemoveAt(i);
                        }
                    }

                    //clear all SDI window references
                    m_dictTaskPanes.Clear();
                }
                else
                {
                    //going to SDI
                    //if it exists, update the task pane & add it to the list
                    if (Globals.ThisAddIn.CustomTaskPanes.Count == 1)
                    {
                        m_dictTaskPanes.Add((Word.Window)Globals.ThisAddIn.CustomTaskPanes[0].Window, Globals.ThisAddIn.CustomTaskPanes[0]);
                    }
                }

                //switch internal state
                m_fLastSdiMode = !m_fLastSdiMode;

                //update the ribbon
                UpdateButtonState();
            }
        }

        private void ctp_VisibleChange(object ctp, System.EventArgs eventArgs)
        {
            //tell the ribbon to refresh
            UpdateButtonState();
        }

        /// <summary>
        /// Connect the VisibleChanged event for an instance of our task pane.
        /// </summary>
        /// <param name="ctp">A CustomTaskPane specifying the new instance of the task pane.</param>
        public void ConnectTaskPaneEvents(CustomTaskPane ctp)
        {
            ctp.VisibleChanged += new System.EventHandler(ctp_VisibleChange);
        }
    }
}
