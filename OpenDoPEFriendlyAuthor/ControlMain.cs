//Copyright (c) Microsoft Corporation.  All rights reserved.
using System;
using System.Diagnostics;
using Office = Microsoft.Office.Core;
using System.Globalization;

namespace XmlMappingTaskPane.Controls
{
    public partial class ControlMain : Controls.ControlBase
    {
        private bool m_inDrop;

        /// <summary>
        /// An enumeration of the reasons a refresh event might be fired.
        /// </summary>
        internal enum ChangeReason { DocumentChanged, PartAdded, PartDeleted, PartLoaded, NodeAdded, NodeDeleted, NodeReplaced, DragDrop, OnEnter };

        public ControlMain()
        {
            InitializeComponent();
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
        internal void RefreshControls(ChangeReason ehReason, Office.CustomXMLNode mxnOldNode, Office.CustomXMLNode mxnOldParent, Office.CustomXMLNode mxnOldNextSibling, Office.CustomXMLNode mxnNewNode, Office._CustomXMLPart cxpOldPart)
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
                    controlPartList.RefreshPartList(true, true, false, string.Empty, null, null);
                    break;

                case ChangeReason.PartAdded:
                    Debug.WriteLine("UI refresh for stream addition.");
                    controlPartList.RefreshPartList(true, false, false, string.Empty, null, null);
                    break;

                case ChangeReason.PartDeleted:
                    Debug.WriteLine("UI refresh for stream deletion.");
                    Debug.Assert(cxpOldPart != null, "We were handed a NULL cxp?");
                    controlPartList.RefreshPartList(true, false, false, string.Empty, cxpOldPart.Id, null);
                    break;

                case ChangeReason.PartLoaded:
                    Debug.WriteLine("UI refresh for stream load.");
                    controlPartList.RefreshPartList(true, false, false, string.Empty, null, null);
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
                    controlPartList.RefreshPartList(true, false, true, string.Empty, null, null);
                    break;

                case ChangeReason.OnEnter:
                    Debug.WriteLine("UI refresh for BB enter.");
                    Debug.Assert(mxnNewNode != null, "No mxn for CC?");
                    if ((controlTreeView.Options & ControlTreeView.cOptionsAutoSelectNode) != 0)
                    {
                        controlPartList.RefreshPartList(false, false, false, mxnNewNode.OwnerPart.Id, null, mxnNewNode);
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
                controlPartList.EventHandler = value;
                controlTreeView.EventHandler = value;
                controlProperties.EventHandler = value;
            }
        }
    }
}
