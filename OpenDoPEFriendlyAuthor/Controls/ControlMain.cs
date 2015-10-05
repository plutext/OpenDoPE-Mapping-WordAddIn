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
using Office = Microsoft.Office.Core;
using System.Globalization;
using OpenDoPEModel;

namespace XmlMappingTaskPane.Controls
{
    public partial class ControlMain : Controls.ControlBase
    {
        private bool m_inDrop;

        public Model model { get; set; }

        /// <summary>
        /// An enumeration of the reasons a refresh event might be fired.
        /// </summary>
        internal enum ChangeReason { DocumentChanged, PartAdded, PartDeleted, PartLoaded, NodeAdded, NodeDeleted, NodeReplaced, DragDrop, OnEnter };

        public Forms.FormSwitchSelectedPart formPartList {get; set;}

        // Whether the mode control buttons are displayed to
        // the user.  This is configurable.
        bool _modeControlEnabled; 
        public bool modeControlEnabled 
        {
            get { return _modeControlEnabled; }
            set { _modeControlEnabled = value; }        
        }


        public ControlMain()
        {
            String modeString = System.Configuration.ConfigurationManager.AppSettings["TaskPane.ModeControlEnabled"];
            _modeControlEnabled = true;
            if (modeString != null)
            {
                Boolean.TryParse(modeString, out _modeControlEnabled);
            }

            InitializeComponent();
            formPartList = new Forms.FormSwitchSelectedPart();
            formPartList.controlPartList.controlMain = this;


        }

        #region Refresh methods

        /// <summary>
        /// Refresh the task pane based on some event.
        /// </summary>
        /// <param name="ehReason">A ChangeReason value specifying the reason for the refresh request.</param>
        /// <param name="mxnOldNode">A CustomXMLNode specifying the deleted node (if applicable).</param>
        /// <param name="mxnOldParent">A CustomXMLNode specifying the former parent node of the deleted node (if applicable).</param>
        /// <param name="mxnOldNextSibling">A CustomXMLNode specifying the former next sibling node of the affected node (if applicable).</param>
        /// <param name="mxnNewNode">A CustomXMLNode specifying the new node (if applicable).</param>
        /// <param name="cxpOldPart">A CustomXMLPart specifying the XML part being deleted (if applicable).</param>
        internal void RefreshControls(ChangeReason ehReason, 
            Office.CustomXMLNode mxnOldNode, Office.CustomXMLNode mxnOldParent, 
            Office.CustomXMLNode mxnOldNextSibling, Office.CustomXMLNode mxnNewNode, 
            Office._CustomXMLPart cxpOldPart)
        {
            // determine why we've been asked to refresh the pane
            switch (ehReason)
            {
                case ChangeReason.DocumentChanged:
                    Debug.WriteLine("UI refresh for document change.");
                    if (!Globals.ThisAddIn.Application.ShowWindowsInTaskbar)
                    {
                        EventHandler.ChangeCurrentDocument();
                    }
                    formPartList.controlPartList.RefreshPartList(true, true, false, string.Empty, null, null);
                    break;

                case ChangeReason.PartAdded:
                    Debug.WriteLine("UI refresh for stream addition.");
                    formPartList.controlPartList.RefreshPartList(true, false, false, string.Empty, null, null);
                    break;

                case ChangeReason.PartDeleted:
                    Debug.WriteLine("UI refresh for stream deletion.");
                    Debug.Assert(cxpOldPart != null, "We were handed a NULL cxp?");
                    formPartList.controlPartList.RefreshPartList(true, false, false, string.Empty, cxpOldPart.Id, null);
                    break;

                case ChangeReason.PartLoaded:
                    Debug.WriteLine("UI refresh for stream load.");
                    formPartList.controlPartList.RefreshPartList(true, false, false, string.Empty, null, null);
                    break;

                case ChangeReason.NodeAdded:
                    Debug.WriteLine("UI refresh for node addition.");
                    controlTreeView.AddMxnToTree(mxnNewNode);
                    break;

                case ChangeReason.NodeDeleted:
                    Debug.WriteLine("UI refresh for node deletion.");
                    controlTreeView.RemoveMxnFromTree(mxnOldNode, mxnOldParent, mxnOldNextSibling, null);
                    break;

                case ChangeReason.NodeReplaced:
                    Debug.WriteLine("UI refresh for node replacement.");
                    controlTreeView.ReplaceMxnInTree(mxnOldNode, mxnNewNode);
                    break;

                case ChangeReason.DragDrop:
                    Debug.WriteLine("UI refresh for drag/drop of XML.");
                    formPartList.controlPartList.RefreshPartList(true, false, true, string.Empty, null, null);
                    break;

                case ChangeReason.OnEnter:
                    Debug.WriteLine("UI refresh for BB enter.");
                    Debug.Assert(mxnNewNode != null, "No mxn for CC?");
                    if ((controlTreeView.Options & ControlTreeView.cOptionsAutoSelectNode) != 0)
                    {
                        formPartList.controlPartList.RefreshPartList(false, false, false, mxnNewNode.OwnerPart.Id, null, mxnNewNode);
                    }
                    break;

                default:
                    Debug.Assert(false, "Unknown refresh event", "We got called on a refresh event that doesn't exist??");
                    break;
            }
        }

        /// <summary>
        /// Refresh the contents of the tree view control.
        /// </summary>
        /// <param name="mxnNodeToSelect">Optional. A CustomXMLNode specifying a node which should be selected after the refresh is completed.</param>
        internal void RefreshTreeControl(Office.CustomXMLNode mxnNodeToSelect)
        {
            controlTreeView.RefreshTree(mxnNodeToSelect);
        }

        /// <summary>
        /// Select a node within the tree view.
        /// </summary>
        /// <param name="mxnToSelect">A CustomXMLNode specifying a node which should be selected after the refresh is completed.</param>
        internal void SelectNodeFromTree(Office.CustomXMLNode mxnToSelect)
        {
            //select the node (happens on the calling thread)
            controlTreeView.SelectNodeFromTree(mxnToSelect);
        }

        /// <summary>
        /// Refresh the contents of the property grid.
        /// </summary>
        /// <param name="mxn">A CustomXMLNode specifying the node whose properties we want to display.</param>
        internal void RefreshProperties(Office.CustomXMLNode mxn)
        {
            controlProperties.RefreshProperties(mxn);
        }

        /// <summary>
        /// Show the XPath for a content control which is mapped, but dangling.
        /// </summary>
        /// <param name="xpath"></param>
        internal void WarnViaProperties(string xpath)
        {
            controlProperties.XPathWarning(xpath);
        }

        internal void PropertiesClear()
        {
            controlProperties.clear();
        }

        /// <summary>
        /// Refresh the task pane based on a change in the settings.
        /// </summary>
        /// <param name="newOptions">An integer specifying the updated set of settings.</param>
        internal void RefreshSettings(int newOptions)
        {
            if (controlTreeView.Options != newOptions)
            {
                // set visibility of property page
                if ((newOptions & ControlTreeView.cOptionsShowPropertyPage) == 0)
                {
                    HidePropertyPage();
                }
                else if (this.Height >= 400)
                {
                    ShowPropertyPage();
                }

                // refresh tree iff a relevant setting changed
                if ((controlTreeView.Options ^ newOptions) != ControlTreeView.cOptionsShowPropertyPage && (controlTreeView.Options ^ newOptions) != ControlTreeView.cOptionsAutoSelectNode && (controlTreeView.Options ^ newOptions) != (ControlTreeView.cOptionsAutoSelectNode + ControlTreeView.cOptionsShowPropertyPage))
                {
                    controlTreeView.RefreshTree(null);
                }

                // set the new options on the treeview
                controlTreeView.Options = newOptions;
            }
        }

        #endregion

        #region Property Grid visibility

        /// <summary>
        /// Show the property grid within the task pane.
        /// </summary>
        private void ShowPropertyPage()
        {
            splitContainer.Panel2Collapsed = false;
        }

        /// <summary>
        /// Hide the property grid within the task pane.
        /// </summary>
        private void HidePropertyPage()
        {
            splitContainer.Panel2Collapsed = true;
        }

        private void ControlMain_Resize(object sender, EventArgs e)
        {
            //hide the property page when we get small
            if (this.Height < 400)
            {
                HidePropertyPage();
            }
            else if ((controlTreeView.Options & ControlTreeView.cOptionsShowPropertyPage) != 0)
            {
                ShowPropertyPage();
            }
        }

        #endregion

        /// <summary>
        /// Notify ourselves that a drag/drop occurred, so we can handle it if it was dropped in Word.
        /// </summary>
        internal void NotifyDragDrop(bool active)
        {
            //set a flag indicating whether we're in drag/drop
            m_inDrop = active;

            if (float.Parse(CurrentDocument.Application.Version, CultureInfo.InvariantCulture) > 12)
            {
                //TODO: start a custom undo record...
            }
        }

        /// <summary>
        /// True if a drag/drop happened recently, False otherwise.
        /// </summary>
        internal bool RecentDragDrop
        {
            get
            {
                return m_inDrop;
            }
        }


        /// <summary>
        /// Set a new DocumentEvents object on this control and its children.
        /// </summary>
        internal DocumentEvents EventHandlerAndOnChildren
        {
            set
            {
                EventHandler = value;
                formPartList.controlPartList.EventHandler = value;
                controlTreeView.EventHandler = value;
                controlProperties.EventHandler = value;
            }

            get
            {
                return EventHandler;
            }
        }

        private void controlTreeView_Load(object sender, EventArgs e)
        {

        }
    }
}
