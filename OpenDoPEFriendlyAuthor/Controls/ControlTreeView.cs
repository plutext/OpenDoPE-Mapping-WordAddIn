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

namespace XmlMappingTaskPane.Controls
{
    public partial class ControlTreeView : Controls.ControlBase
    {
        static Logger log = LogManager.GetLogger("ControlTreeView");

        // the current option set
        private int m_grfOptions;

        /// <summary>
        /// Constant values corresponding to each option in the bitflag (for comparison).
        /// </summary>
        internal const int cOptionsShowAttributes = 1;
        internal const int cOptionsShowText = 2;
        internal const int cOptionsShowPI = 4;
        internal const int cOptionsShowComments = 8;
        internal const int cOptionsShowPropertyPage = 16;
        internal const int cOptionsAutoSelectNode = 32;

        //update the progress UI every N nodes
        private const int cintUpdateFrequency = 33; 

        private TreeNode m_tnRoot; // the root treenode
        private IDictionary<XmlNode, TreeNode> m_dicXnTn = new Dictionary<XmlNode, TreeNode>();

        // Whether the user can edit the XML tree
        bool _XmlTreeIsEditable;
        public bool XmlTreeIsEditable
        {
            get { return _XmlTreeIsEditable; }
            set { _XmlTreeIsEditable = value; }
        }
        /// <summary>
        /// If set to false, the control type added  via right click will be a text control. 
        /// 
        /// If set to true, user can choose between text, date, drop down list, picture, combo box.
        /// </summary>
        bool _BindControlTypeChoice;
        public bool BindControlTypeChoice
        {
            get { return _BindControlTypeChoice; }
            set { _BindControlTypeChoice = value; }
        }

        /// <summary>
        /// Whether rich text controls are used in favour of picture content controls.
        /// </summary>
        bool _PictureContentControlsReplace = true;


        /// <summary>
        ///  Whether you can right click to add a condition.
        /// </summary>
        bool conditionViaRightClick = true;

        /// <summary>
        /// Helper class which contains OpenDoPE magic for dragging XML elements from the task pane
        /// </summary>
        OpenDopeDragHandler openDopeDragHandler = new OpenDopeDragHandler();

        /// <summary>
        /// Helper class which contains OpenDoPE magic for right clicking on XML elements in the task pane
        /// </summary>
        OpenDopeRightClickHandler openDopeRightClickHandler = new OpenDopeRightClickHandler();

        /// <summary>
        /// Helper class to create a content control mapped to the selected XML node.
        /// </summary>
        OpenDopeCreateMappedControl openDopeCreateMappedControl = new OpenDopeCreateMappedControl();

        public ControlTreeView()
        {
            String xmlTreeEd = System.Configuration.ConfigurationManager.AppSettings["TaskPane.XmlTreeIsEditable"];
            _XmlTreeIsEditable = true;
            if (xmlTreeEd != null)
            {
                Boolean.TryParse(xmlTreeEd, out _XmlTreeIsEditable);
            }

            // We'll use this value as we init the component
            String bindChoice = System.Configuration.ConfigurationManager.AppSettings["TaskPane.BindControlTypeChoice"];
            _BindControlTypeChoice = true;
            if (bindChoice != null)
            {
                Boolean.TryParse(bindChoice, out _BindControlTypeChoice);
                log.Debug("_BindControlTypeChoice: " + _BindControlTypeChoice);
            }

            String picSetting = System.Configuration.ConfigurationManager.AppSettings["ContentControl.Picture.RichText.Override"];
            if (picSetting != null)
            {
                Boolean.TryParse(picSetting, out _PictureContentControlsReplace);
            }

            
            InitializeComponent();

            string conditionViaRightClickStr = System.Configuration.ConfigurationManager.AppSettings["TaskPane.ConditionViaRightClick"];
            if (conditionViaRightClickStr != null
                && conditionViaRightClickStr.ToLower().Equals("false"))
            {
                conditionViaRightClick = false;
            }

            //set up the options
            try
            {
                Options = (int)Registry.CurrentUser.OpenSubKey(System.Configuration.ConfigurationManager.AppSettings["Registry.CurrentUser.SubKey"]).GetValue("Options");
            }
            catch (NullReferenceException nrex)
            {
                Debug.Fail("regkey corruption", "either the user manually deleted the regkeys, or something bad happened." + Environment.NewLine + nrex.Message);
            }
        }

        /// <summary>
        /// Get/set the options specified for the current treeview.
        /// </summary>
        internal int Options
        {
            get
            {
                return m_grfOptions;
            }
            set
            {
                m_grfOptions = value;
            }
        }

        /// <summary>
        /// Get the XmlDocument for the current XML part.
        /// </summary>
        private XmlDocument OwnerDocument
        {
            get
            {
                XmlDocument xdoc = null;
                if (((XmlNode)m_tnRoot.Tag).NodeType == XmlNodeType.Document)
                    xdoc = (XmlDocument)m_tnRoot.Tag;
                else
                    xdoc = ((XmlNode)m_tnRoot.Tag).OwnerDocument;
                return xdoc;
            }
        }

        #region Timer methods

        /// <summary>
        /// Start the loading timer.
        /// </summary>
        private delegate void StartTimerDelegate();
        private void StartTimer()
        {
            if (InvokeRequired)
            {
                Invoke(new StartTimerDelegate(StartTimer), new object[] { });
                return;
            }

            timerLoading.Start();
            labelLoading.Visible = true;
        }

        /// <summary>
        /// Stop the loading timer.
        /// </summary>
        private delegate void StopTimerDelegate();
        private void StopTimer()
        {
            if (InvokeRequired)
            {
                Invoke(new StopTimerDelegate(StopTimer), new object[] { });
                return;
            }

            timerLoading.Stop();
        }

        /// <summary>
        /// Reset the loading timer.
        /// </summary>
        private delegate void ResetTimerDelegate();
        private void ResetTimer()
        {
            if (InvokeRequired)
            {
                Invoke(new ResetTimerDelegate(ResetTimer), new object[] { });
                return;
            }

            timerLoading.Stop();
            labelLoading.Text = string.Empty;
            labelLoading.Visible = false;
        }

        private void timerLoading_Tick(object sender, EventArgs e)
        {
            Debug.WriteLine("timer ticked");
            SetLabelText(Properties.Resources.LoadingMessage);
            StopTimer();
        }

        #endregion

        #region TreeView manipulation methods

        /// <summary>
        /// Select the treeview node corresponding to a specific CustomXMLNode object.
        /// </summary>
        /// <param name="mxnBB">A CustomXMLNode specifying the node to be selected.</param>
        internal delegate void SelectNodeFromTreeDelegate(Office.CustomXMLNode mxnBB);
        internal void SelectNodeFromTree(Office.CustomXMLNode mxnBB)
        {
            if (backgroundWorkerBuildTree.CancellationPending)
            {
                throw new QuitNowException();
            }
            else if (InvokeRequired)
            {
                Invoke(new SelectNodeFromTreeDelegate(SelectNodeFromTree), new object[] { mxnBB });
                return;
            }

            //we select the root (2nd) node to get the DOM
            Debug.Assert(m_tnRoot != null, "Node doesn't exist", "You can't expect a root node?");

            XmlDocument xdoc = OwnerDocument;

            //then go mxn -> tn
            XmlNode xnBB = Utilities.XnFromMxn(xdoc, mxnBB, null);
            TreeNode tnBB = m_dicXnTn[xnBB];
            treeView.SelectedNode = tnBB;

            //update the properties
            ControlMain controlMain = GetMainControl();
            controlMain.RefreshProperties(mxnBB);
        }

        public void DeselectNode()
        {
            treeView.SelectedNode = null;
        }

        /// <summary>
        /// Clear out all of the tree nodes from a treeview.
        /// </summary>
        /// <param name="tv">A TreeView whose contents should be cleared.</param>
        private delegate void ClearTreeNodesDelegate(TreeView tv);
        private void ClearTreeNodes(TreeView tv)
        {
            if (InvokeRequired)
            {
                Invoke(new ClearTreeNodesDelegate(ClearTreeNodes), new object[] { tv });
                return;
            }

            tv.Nodes.Clear();
        }

        /// <summary>
        /// Set the label on the loading control.
        /// </summary>
        /// <param name="s">A string specifying the new text for the control.</param>
        private delegate void SetLabelTextDelegate(string s);
        private void SetLabelText(string s)
        {
            if (InvokeRequired)
            {
                Invoke(new SetLabelTextDelegate(SetLabelText), new object[] { s });
                return;
            }

            labelLoading.Text = s;
            labelLoading.PerformLayout(this, "Text");
        }

        /// <summary>
        /// Expand a specific node in the tree.
        /// </summary>
        /// <param name="tn">A TreeNode specifying the node to expand.</param>
        private delegate void ExpandTreeNodeDelegate(TreeNode tn);
        private void ExpandTreeNode(TreeNode tn)
        {
            if (backgroundWorkerBuildTree.CancellationPending)
            {
                throw new QuitNowException();
            }
            else if (InvokeRequired)
            {
                Invoke(new ExpandTreeNodeDelegate(ExpandTreeNode), new object[] { tn });
                return;
            }

            tn.Expand();
        }

        /// <summary>
        /// Ensure a specific tree element is visible.
        /// </summary>
        /// <param name="tn">A TreeNode specifying the node to make visible.</param>
        private delegate void EnsureTreeNodeDelegate(TreeNode tn);
        private void EnsureTreeNode(TreeNode tn)
        {
            if (backgroundWorkerBuildTree.CancellationPending)
            {
                throw new QuitNowException();
            }
            else if (InvokeRequired)
            {
                Invoke(new EnsureTreeNodeDelegate(EnsureTreeNode), new object[] { tn });
                return;
            }

            tn.EnsureVisible();
        }

        /// <summary>
        /// Add a tree node to the control.
        /// </summary>
        /// <param name="tnBase">The TreeNode within which the new node should be added.</param>
        /// <param name="tnToAdd">The TreeNode to be added to the control.</param>
        private delegate void AddTreeNodeDelegate(TreeNode tnBase, TreeNode tnToAdd);
        private void AddTreeNode(TreeNode tnBase, TreeNode tnToAdd)
        {
            if (backgroundWorkerBuildTree.CancellationPending)
            {
                throw new QuitNowException();
            }
            else if (InvokeRequired)
            {
                Invoke(new AddTreeNodeDelegate(AddTreeNode), new object[] { tnBase, tnToAdd });
                return;
            }

            tnBase.Nodes.Add(tnToAdd);
        }

        /// <summary>
        /// Add the root node to the tree view control.
        /// </summary>
        /// <param name="tnToAdd">The TreeNode to be added to the control.</param>
        private delegate void AddRootTreeNodeDelegate(TreeNode tnToAdd);
        private void AddRootTreeNode(TreeNode tnToAdd)
        {
            if (backgroundWorkerBuildTree.CancellationPending)
            {
                throw new QuitNowException();
            }
            else if (InvokeRequired)
            {
                Invoke(new AddRootTreeNodeDelegate(AddRootTreeNode), new object[] { tnToAdd });
                return;
            }

            treeView.Nodes.Add(tnToAdd);
        }

        /// <summary>
        /// Set the image associated with a tree node.
        /// </summary>
        /// <param name="tn">The TreeNode whose image should be changed.</param>
        /// <param name="ImageIndex">An integer containing the index of the desired image in the image collection.</param>
        private delegate void SetTreeNodeImageDelegate(TreeNode tn, int ImageIndex);
        private void SetTreeNodeImage(TreeNode tn, int ImageIndex)
        {
            if (backgroundWorkerBuildTree.CancellationPending)
            {
                throw new QuitNowException();
            }
            else if (InvokeRequired)
            {
                Invoke(new SetTreeNodeImageDelegate(SetTreeNodeImage), new object[] { tn, ImageIndex });
                return;
            }

            tn.ImageIndex = ImageIndex;
            tn.SelectedImageIndex = ImageIndex;
        }

        /// <summary>
        /// Add a tree node to the tree view control at a specific position.
        /// </summary>
        /// <param name="tnBase">A TreeNode specifying the parent of the new tree node.</param>
        /// <param name="tnToInsert">A TreeNode specifying the tree node to be added.</param>
        /// <param name="Position">An integer specifying the ordinal position of the new node within the parent's child nodes.</param>
        private delegate void InsertTreeNodeDelegate(TreeNode tnBase, TreeNode tnToInsert, int Position);
        private void InsertTreeNode(TreeNode tnBase, TreeNode tnToInsert, int Position)
        {
            if (backgroundWorkerBuildTree.CancellationPending)
            {
                throw new QuitNowException();
            }
            else if (InvokeRequired)
            {
                Invoke(new InsertTreeNodeDelegate(InsertTreeNode), new object[] { tnBase, tnToInsert, Position });
                return;
            }

            tnBase.Nodes.Insert(Position, tnToInsert);
        }

        #endregion

        #region Tree Refresh methods/events

        /// <summary>
        /// Refresh the tree view.
        /// </summary>
        /// <param name="mxnToSelect">A CustomXMLNode specifying the node which should be selected after refresh completes.</param>
        internal void RefreshTree(Office.CustomXMLNode mxnToSelect)
        {
            // build the tree view on a background thread
            //  this ensures that a large CustomXMLPart cannot hang the application, and 
            //  that we can cancel load if the user wishes to
            if (!backgroundWorkerMain.IsBusy)
            {
                backgroundWorkerMain.RunWorkerAsync(mxnToSelect);  

                // Or could do
                //TreeViewElements tve = new TreeViewElements(CurrentPart, mxnToSelect);
                //backgroundWorkerBuildTree.RunWorkerAsync(tve);
            }
            else
            {
                //Debug.Fail("unable to execute a refresh because the worker was already busy.");
                log.Warn("unable to execute a refresh because the worker was already busy.");
                // 2012 09 29 .. just silently ignore this.
                // Hopefully the worker will finish in due course,
                // and a later refresh will succeed.
            }
        }

        /// <summary>
        /// Tell the background worker thread to populate the tree view.
        /// </summary>
        private void backgroundWorkerMain_DoWork(object sender, DoWorkEventArgs e)
        {
        
        LWait:

            //check if we've got a refresh going
            if (backgroundWorkerBuildTree.IsBusy)
            {
                //we're already refreshing, so cancel it
                backgroundWorkerBuildTree.CancelAsync();

                //wait for it...
                while (backgroundWorkerBuildTree.IsBusy)
                {
                    Thread.Sleep(0);
                }
            }

            //start a timer to ensure we only show progress if necessary
            StartTimer();

            //populate it
            if (CurrentPart.DocumentElement != null)
            {
                TreeViewElements tve = new TreeViewElements(CurrentPart, (Office.CustomXMLNode)e.Argument);
                if (!backgroundWorkerBuildTree.IsBusy)
                {
                    backgroundWorkerBuildTree.RunWorkerAsync(tve);
                }
                else
                    goto LWait; //we started a task between now and the last cancel
            }
            else
            {
                //clear out
                ClearTreeNodes(treeView);
                ResetTimer();                
            }
        }

        /// <summary>
        /// Do the work to actually build the tree view.
        /// </summary>
        private void backgroundWorkerBuildTree_DoWork(object sender, DoWorkEventArgs e)
        {
            log.Debug("fired");

            try
            {
                //load up the stream into an XML DOM
                TreeViewElements tve = (TreeViewElements)e.Argument;

                //load the DOM
                tve.TreeViewDOM.PreserveWhitespace = true; //keep whitespace, like Office does
                tve.TreeViewDOM.XmlResolver = null; // no entity/DTD expansion
                tve.TreeViewDOM.LoadXml(tve.TreeViewPart.XML);

                //get XSDs
                if (CurrentPart.SchemaCollection != null && CurrentPart.SchemaCollection.Count > 0)
                {
                    ValidationEventHandler veh = new ValidationEventHandler(xs_veh);
                    XmlSchemaSet xss = new XmlSchemaSet();
                    foreach (Office.CustomXMLSchema cxs in CurrentPart.SchemaCollection)
                    {
                        if (!string.IsNullOrEmpty(cxs.NamespaceURI) && !string.IsNullOrEmpty(cxs.Location))
                            xss.Add(cxs.NamespaceURI, cxs.Location);
                    }

                    if (xss.Count != 0)
                    {
                        xss.Compile();
                        tve.TreeViewDOM.Schemas = xss;
                        tve.TreeViewDOM.Validate(veh);
                    }
                }

                //clear the hashtable
                m_dicXnTn.Clear();

                //get total # of nodes
                int intNumberProcessed = 0;
                int intNumberTotal = tve.TreeViewDOM.SelectNodes("//*").Count;
                if ((Options & cOptionsShowAttributes) != 0)
                    intNumberTotal += (tve.TreeViewDOM.SelectNodes("//@*").Count * 2); //2x because we have attribute and its text node
                if ((Options & cOptionsShowComments) != 0)
                    intNumberTotal += tve.TreeViewDOM.SelectNodes("//comment()").Count;
                if ((Options & cOptionsShowPI) != 0)
                    intNumberTotal += tve.TreeViewDOM.SelectNodes("//processing-instruction()").Count;
                if ((Options & cOptionsShowText) != 0)
                    intNumberTotal += tve.TreeViewDOM.SelectNodes("//text()").Count;

                //if too few, never start a timer
                if (intNumberTotal < 100)
                    ResetTimer();

                //report progress
                backgroundWorkerBuildTree.ReportProgress(0, intNumberTotal);

                XmlNode xn = null;
                if ((tve.TreeViewDOM.DocumentElement.PreviousSibling != null && tve.TreeViewDOM.DocumentElement.PreviousSibling.NodeType != XmlNodeType.XmlDeclaration) || tve.TreeViewDOM.DocumentElement.NextSibling != null)
                    xn = tve.TreeViewDOM as XmlNode;
                else
                    xn = tve.TreeViewDOM.DocumentElement;

                //build a tree node for the root
                m_tnRoot = TnBuildFromXn(xn);

                //clear the tree - invoking on another thread
                ClearTreeNodes(treeView);

                //start at the root and populate
                AddRootTreeNode(m_tnRoot);
                m_tnRoot.Tag = xn;

                //work through all nodes
                PopulateTreeNode(xn, m_tnRoot, intNumberProcessed, intNumberTotal);

                //make sure the first level is expanded
                ExpandTreeNode(m_tnRoot);
                if (((XmlNode)m_tnRoot.Tag).NodeType == XmlNodeType.Document)
                    foreach (TreeNode tnNode in m_tnRoot.Nodes)
                    {
                        ExpandTreeNode(tnNode);
                    }

                //make sure we're at the top
                if (tve.NodeToSelect == null)
                    EnsureTreeNode(m_tnRoot);
                else
                    SelectNodeFromTree(tve.NodeToSelect);

                //turn off the loading screen
                ResetTimer();
            }
            catch (QuitNowException qnex)
            {
                Debug.WriteLine(qnex.Message);
                //clear up, and quit
                ClearTreeNodes(treeView);
                e.Cancel = true;
                return;
            }
        }

        /// <summary>
        /// Report progress on building the tree view (by updating the label on the tree view appropriately).
        /// </summary>
        private void backgroundWorkerBuildTree_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            SetLabelText(string.Format(CultureInfo.CurrentCulture, Properties.Resources.LoadingMessageWithPercentage, (int)(((double)e.ProgressPercentage / (int)e.UserState) * 100)));
        }

        private void xs_veh(object o, ValidationEventArgs vea)
        {
            //we don't care about the validation errors, we just want to validate part against the schema set and know if it was successful or not
        }

        #endregion

        #region External add/remove/replace methods/events

        /// <summary>
        /// Add a node that was added to the XML part.
        /// </summary>
        /// <param name="mxnNew">The CustomXMLNode used to create the coresponding tree node.</param>
        internal void AddMxnToTree(Office.CustomXMLNode mxnNew)
        {
            //if background worker is building the tree, block on that call
            while (backgroundWorkerBuildTree.IsBusy)
            {
                Debug.WriteLine("waiting for tree to be ready so we can add a node...");
                Thread.Sleep(500);
            }

            //get the xn for the parent
            Debug.Assert(mxnNew.ParentNode != null, "null parentNode", "how can a node not have a parent?!!?");
            XmlNode xnParent = Utilities.XnFromMxn(OwnerDocument, mxnNew.ParentNode, null);

            //get the xn for the next sibling
            XmlNode xnNextSibling = null;
            if (mxnNew.NextSibling != null)
                xnNextSibling = Utilities.XnFromMxn(OwnerDocument, mxnNew.NextSibling, mxnNew);

            //build an XmlNode
            XmlNode xnNew = Utilities.XnBuildFromMxn(mxnNew, OwnerDocument);

            //add the node into the DOM
            if (xnNew.NodeType == XmlNodeType.Attribute)
                xnNew = xnParent.Attributes.Append((XmlAttribute)xnNew);
            else
                xnNew = xnParent.InsertBefore(xnNew, xnNextSibling);
            xnNew.OwnerDocument.Normalize();

            //normalizing the DOM disconnects the text node(?!), so get it back
            if (xnNew.NodeType != XmlNodeType.Attribute)
            {
                if (xnNextSibling == null)
                {
                    xnNew = xnParent.LastChild;
                }
                else
                {
                    xnNew = xnNextSibling.PreviousSibling;
                }
            }

            xnNew = AddNodeToTree(xnParent, xnNextSibling, xnNew);
        }

        /// <summary>
        /// Remove a node that was removed from the XML part.
        /// </summary>
        /// <param name="mxnOldNode">The CustomXMLNode removed from the XML part.</param>
        /// <param name="mxnOldParent">The CustomXMLNode specifying the parent of the deleted node.</param>
        /// <param name="mxnOldNextSibling">The CustomXMLNode specifying the next sibling of the deleted node.</param>
        /// <param name="mxnNew">A CustomXMLNode specifying the node the deleted node is being replaced with. Only used during replace actions.</param>
        internal void RemoveMxnFromTree(Office.CustomXMLNode mxnOldNode, Office.CustomXMLNode mxnOldParent, Office.CustomXMLNode mxnOldNextSibling, Office.CustomXMLNode mxnNew)
        {
            //if background worker is building the tree, block on that call
            while (backgroundWorkerBuildTree.IsBusy)
            {
                Debug.WriteLine("waiting for tree to be ready so we can add a node...");
                Thread.Sleep(500);
            }

            //assume we have a parent
            Debug.Assert(mxnOldParent != null, "deleted a node without a parent?");

            //hold onto the parent treenode
            TreeNode tnParent = null;
            XmlNode xnParent = null;

            //if we have a next sibling, get it's previous sibling in the local DOM
            if (mxnOldNextSibling != null)
            {
                xnParent = Utilities.XnFromMxn(OwnerDocument, mxnOldParent, null);

                //get the tree node and xml node
                XmlNode xnNextSibling = Utilities.XnFromMxn(OwnerDocument, mxnOldNextSibling, mxnNew);
                Debug.Assert(xnNextSibling.PreviousSibling != null, "asked to delete a prev sibling that doesn't exist");

                //check if this node was in the tree
                if (TnVisibiltyFetchFromGrf(xnNextSibling.PreviousSibling, Options))
                {
                    TreeNode tnOld = m_dicXnTn[xnNextSibling.PreviousSibling];
                    tnParent = tnOld.Parent;

                    //remove from the tree
                    tnOld.Remove();

                    //remove from the list
                    m_dicXnTn.Remove(xnNextSibling.PreviousSibling);
                }

                //remove from the DOM
                Debug.Assert(xnNextSibling.PreviousSibling.LocalName == mxnOldNode.BaseName && xnNextSibling.PreviousSibling.NamespaceURI == mxnOldNode.NamespaceURI, "DOMs falling out of sync - we're removing a different node than the MXSI just did.");
                xnNextSibling.ParentNode.RemoveChild(xnNextSibling.PreviousSibling);
            }
            else if (mxnOldNode.NodeType == Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute)
            {
                //find the attribute with the same namespace and basename on the parent and kill it
                xnParent = Utilities.XnFromMxn(OwnerDocument, mxnOldParent, null);

                XmlAttribute xaToRemove = null;
                foreach (XmlAttribute xaAttribute in xnParent.Attributes)
                    if (xaAttribute.NamespaceURI == mxnOldNode.NamespaceURI && xaAttribute.LocalName == mxnOldNode.BaseName)
                    {
                        if (TnVisibiltyFetchFromGrf(xaAttribute, Options) == true)
                        {
                            TreeNode tnOld = m_dicXnTn[xaAttribute];
                            tnParent = tnOld.Parent;

                            //remove from the tree
                            tnOld.Remove();

                            //remove from the list
                            m_dicXnTn.Remove(xaAttribute);

                            //cache it to remove later
                            xaToRemove = xaAttribute;
                        }
                    }

                //remove it
                Debug.Assert(xaToRemove != null, "didn't find the right attr?");
                xnParent.Attributes.Remove(xaToRemove);
            }
            else
            {
                xnParent = Utilities.XnFromMxn(OwnerDocument, mxnOldParent, null);

                Debug.Assert(xnParent.LastChild != null || (!string.IsNullOrEmpty(mxnOldNode.NodeValue) && mxnOldNode.NodeType == Office.MsoCustomXMLNodeType.msoCustomXMLNodeText), "why don't we have this child node?");
                if (xnParent.LastChild != null)
                {
                    if (TnVisibiltyFetchFromGrf(xnParent.LastChild, Options))
                    {
                        TreeNode tnOld;
                        Debug.Assert(m_dicXnTn.ContainsKey(xnParent.LastChild) || (xnParent.LastChild.NodeType == XmlNodeType.Text && string.IsNullOrEmpty(xnParent.LastChild.InnerText)), "why don't we have this text node?");
                        if (m_dicXnTn.ContainsKey(xnParent.LastChild))
                        {
                            tnOld = m_dicXnTn[xnParent.LastChild];
                            tnParent = tnOld.Parent;

                            //remove the tree node
                            tnOld.Remove();

                            //remove from the list
                            m_dicXnTn.Remove(xnParent.LastChild);
                        }
                    }

                    Debug.Assert(mxnOldNode.BaseName == xnParent.LastChild.LocalName && mxnOldNode.NamespaceURI == xnParent.LastChild.NamespaceURI, "we're deleting the wrong node");
                    xnParent.RemoveChild(xnParent.LastChild);
                }
            }

            //fix the parent's icon
            if (tnParent != null)
            {
                bool bIsLeafNode = IsLeafNode(tnParent);

                if (!bIsLeafNode)
                {
                    tnParent.ImageIndex = 0;
                    tnParent.SelectedImageIndex = 0;
                }
                else
                {
                    tnParent.ImageIndex = 2;
                    tnParent.SelectedImageIndex = 2;
                }
            }
        }

        /// <summary>
        /// Replace a node that was replaced in the XML part.
        /// </summary>
        /// <param name="mxnOld">The CustomXMLNode removed from the XML part.</param>
        /// <param name="mxnNew">The CustomXMLNode added to the XML part.</param>
        internal void ReplaceMxnInTree(Office.CustomXMLNode mxnOld, Office.CustomXMLNode mxnNew)
        {
            //if background worker is building the tree, block on that call
            while (backgroundWorkerBuildTree.IsBusy)
            {
                Debug.WriteLine("waiting for tree to be ready so we can add a node...");
                Thread.Sleep(500);
            }

            XmlNode xn = null;            
            if(mxnNew.PreviousSibling != null)
            {
                xn = Utilities.XnFromMxn(OwnerDocument, mxnNew.PreviousSibling, null);
                xn = xn.NextSibling;
                
            }
            else
            {
                xn = Utilities.XnFromMxn(OwnerDocument, mxnNew.ParentNode, null);
                if (mxnOld.NodeType == Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute)
                    xn = xn.Attributes[mxnOld.BaseName, mxnOld.NamespaceURI];
                else
                    xn = xn.FirstChild;
            }

            TreeNode tn = null;
            bool fExpandTreeNode = false;
            if (TnVisibiltyFetchFromGrf(xn, Options))
            {
                tn = m_dicXnTn[xn];
                fExpandTreeNode = tn.IsExpanded;
            }

            //replace by deleting and adding from our DOM
            RemoveMxnFromTree(mxnOld, mxnNew.ParentNode, mxnNew.NextSibling, mxnNew);
            AddMxnToTree(mxnNew);

            if (tn != null && fExpandTreeNode)
            {
                xn = Utilities.XnFromMxn(OwnerDocument, mxnNew, null);
                tn = m_dicXnTn[xn];
                tn.Expand();
            }
        }

        /// <summary>
        /// Add a new node to our local XML tree.
        /// </summary>
        /// <param name="xnParent">An XmlNode specifying the parent of the new node.</param>
        /// <param name="xnNextSibling">An XmlNode specifying the next sibling of the new node. Specify null if this should be the last child of the parent node.</param>
        /// <param name="xnNew">An XmlNode specifying the node to be added.</param>
        /// <returns>An XmlNode specifying the node added to our local XML tree.</returns>
        private XmlNode AddNodeToTree(XmlNode xnParent, XmlNode xnNextSibling, XmlNode xnNew)
        {
            //create a treenode
            TreeNode tnToAdd = TnBuildFromXn(xnNew);

            //add it to the tree
            //unless we get back that we don't care about this node type 
            //(e.g. they turned it off in the properties), then exit
            if (tnToAdd != null)
            {
                AddXnToTree(xnNew, xnParent, xnNextSibling);
            }

            return xnNew;
        }

        /// <summary>
        /// Add a node to the tree view.
        /// </summary>
        /// <param name="xnNew">An XmlNode specifying the node to be added.</param>
        /// <param name="xnParent">An XmlNode specifying the parent of the new node.</param>
        /// <param name="xnNextSibling">An XmlNode specifying the next sibling of the new node. Specify null if this should be the last child of the parent node.</param>
        private void AddXnToTree(XmlNode xnNew, XmlNode xnParent, XmlNode xnNextSibling)
        {
            TreeNode tnNew = m_dicXnTn[xnNew];
            TreeNode tnParent = m_dicXnTn[xnParent];
            TreeNode tnNextSibling = null;
            if (xnNextSibling != null)
            {
                tnNextSibling = m_dicXnTn[xnNextSibling];
            }

            if (xnNew.NodeType == XmlNodeType.Attribute)
            {
                //add it within the attributes
                if (tnNextSibling != null)
                {
                    InsertTreeNode(tnParent, tnNew, tnNextSibling.Index);
                    PopulateTreeNode(xnNew, tnNew, 0, 1);
                }
                //no, add it within the elements
                else
                {
                    if (tnParent.Nodes.Count > 0)
                    {
                        foreach (TreeNode tn in tnParent.Nodes)
                        {
                            //find the first non-attribute or attribute sorted after this one, put it before that tree node
                            if (((XmlNode)tn.Tag).NodeType != XmlNodeType.Attribute || (((XmlNode)tn.Tag).NodeType == XmlNodeType.Attribute && string.Compare(tn.Text, tnNew.Text, StringComparison.CurrentCulture) > 0) /*&& (tn.PrevNode == null || ((XmlNode)tn.PrevNode.Tag).NodeType == XmlNodeType.Attribute)*/)
                            {
                                InsertTreeNode(tnParent, tnNew, tn.Index);
                                PopulateTreeNode(xnNew, tnNew, 0, 1);
                                break;
                            }
                        }
                    }
                    else
                    {
                        //it's the only child node, so just put it at the end
                        AddTreeNode(tnParent, tnNew);
                        PopulateTreeNode(xnNew, tnNew, 0, 1);
                    }
                }
            }
            else
            {
                //add it within the child nodes				
                if (tnNextSibling != null)
                {
                    InsertTreeNode(tnParent, tnNew, tnNextSibling.Index);
                    PopulateTreeNode(xnNew, tnNew, 0, 1);
                }
                else
                {
                    InsertTreeNode(tnParent, tnNew, tnParent.Nodes.Count);
                    PopulateTreeNode(xnNew, tnNew, 0, 1);
                }
            }

            //fix the parent's icon
            if (xnNew.NodeType == XmlNodeType.Element)
            {
                SetTreeNodeImage(tnParent, 0);
            }
        }

        /// <summary>
        /// Populate the tree view based on the XML structure of a node.
        /// </summary>
        /// <param name="xn">The XmlNode to use to populate the tree.</param>
        /// <param name="tn">The TreeNode to populate.</param>
        /// <param name="NumberProcessed">An integer specifying the number of nodes already processed in the XML tree.</param>
        /// <param name="NumberTotal">An integer specifying the total number of node in the XML tree.</param>
        /// <returns></returns>
        private int PopulateTreeNode(XmlNode xn, TreeNode tn, int NumberProcessed, int NumberTotal)
        {
            int i = NumberProcessed;

            //create attributes
            if (xn.Attributes != null)
            {
                foreach (XmlNode xnChild in xn.Attributes)
                {
                    TreeNode tnChild = TnBuildFromXn(xnChild);
                    if (tnChild != null)
                    {
                        if (tn.Nodes.Count > 0)
                        {
                            //check the existing list first
                            bool placed = false;
                            foreach (TreeNode tnAttr in tn.Nodes)
                            {
                                //find the first non-attribute or attribute sorted after this one, put it before that tree node
                                if (string.Compare(tnAttr.Text, tnChild.Text, StringComparison.CurrentCulture) > 0)
                                {
                                    InsertTreeNode(tn, tnChild, tnAttr.Index);
                                    placed = true;
                                    break;
                                }
                            }

                            //if it didn't fall into the existing list, put it at the end
                            if (!placed)
                            {
                                AddTreeNode(tn, tnChild);
                            }
                        }
                        else
                        {
                            //no children yet, put it at the end
                            AddTreeNode(tn, tnChild);
                        }

                        i++;
                        if (NumberTotal > cintUpdateFrequency && i % (int)(NumberTotal / cintUpdateFrequency) == 0 && backgroundWorkerBuildTree.IsBusy)
                            backgroundWorkerBuildTree.ReportProgress(i, NumberTotal);

                        i = PopulateTreeNode(xnChild, tnChild, i, NumberTotal);
                    }
                }
            }

            //create children
            foreach (XmlNode xnChild in xn.ChildNodes)
            {
                TreeNode tnChild = TnBuildFromXn(xnChild);
                if (tnChild != null)
                {
                    AddTreeNode(tn, tnChild);
                    i++;
                    if (NumberTotal > cintUpdateFrequency && i % (int)(NumberTotal / cintUpdateFrequency) == 0 && backgroundWorkerBuildTree.IsBusy)
                        backgroundWorkerBuildTree.ReportProgress(i, NumberTotal);

                    i = PopulateTreeNode(xnChild, tnChild, i, NumberTotal);
                }
            }

            return i;
        }

        /// <summary>
        /// Build a tree node out of an XML node, to put into the tree view control.
        /// </summary>
        /// <param name="xn">The XmlNode from which to create the tree node.</param>
        /// <returns>The corresponding TreeNode.</returns>
        private TreeNode TnBuildFromXn(XmlNode xn)
        {
            string strName = null;
            TreeNode tnResult = null;

            //write out the proper content into the node
            switch (xn.NodeType)
            {
                case XmlNodeType.Attribute:
                    //hide namespace declarations & attribs if not selected in options
                    if (!xn.Prefix.Equals("xmlns") && !xn.Name.Equals("xmlns") && (Options & cOptionsShowAttributes) != 0)
                    {
                        strName = xn.LocalName;
                    }
                    break;

                case XmlNodeType.Element:
                    {
                        strName = xn.LocalName;
                        break;
                    }

                case XmlNodeType.ProcessingInstruction:
                    {
                        if ((Options & cOptionsShowPI) != 0)
                        {
                            XmlProcessingInstruction xpi = xn as XmlProcessingInstruction;
                            strName = "<?" + xpi.Target + " " + xpi.Value + "?>";
                        }
                        break;
                    }

                case XmlNodeType.Text:
                    if (!xn.ParentNode.Prefix.Equals("xmlns") && !xn.ParentNode.Name.Equals("xmlns") && (Options & cOptionsShowText) != 0)
                        strName = xn.Value;
                    break;

                case XmlNodeType.Comment:
                    if ((Options & cOptionsShowComments) != 0)
                        strName = "<!--" + xn.Value + "-->";
                    break;

                case XmlNodeType.Document:
                    strName = "/";
                    break;

                case XmlNodeType.CDATA:
                    strName = "<![CDATA[" + xn.Value + "]]>";
                    break;

                case XmlNodeType.Whitespace:
                case XmlNodeType.SignificantWhitespace:
                    break;
            }

            //if this is something we're displaying, set up the view
            if (!string.IsNullOrEmpty(strName))
            {
                //add to hashtables
                tnResult = new TreeNode(strName);
                m_dicXnTn.Add(xn, tnResult);

                //add the xn to the tn
                tnResult.Tag = xn;

                //put a nice icon on it
                AddIconToTreeNode(xn, tnResult);
            }

            return tnResult;
        }

        #endregion

        #region TreeView events

        private void treeView_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent("FileDrop"))
            {
                bool bNewStream = false;
                StartTimer();
                string[] files = (string[])e.Data.GetData("FileDrop");
                foreach (string file in files)
                {
                    Debug.WriteLine("Attempting to load a file from drag/drop:" + file);

                    //check if this is an xml file
                    XmlDocument xdocNewFile = new XmlDocument();
                    xdocNewFile.XmlResolver = null;
                    try
                    {
                        xdocNewFile.Load(file);

                        //if so, add it to the streams collection
                        try
                        {
                            Object objMissing = Type.Missing;
                            CurrentPartCollection.Add(xdocNewFile.OuterXml, objMissing);
                            bNewStream = true;
                        }
                        catch (COMException ex)
                        {
                            ShowErrorMessage(string.Format(CultureInfo.CurrentCulture, Properties.Resources.ErrorOnPartAdd, ex.Message));
                            ResetTimer();
                        }
                    }
                    catch (XmlException ex)
                    {
                        ShowErrorMessage(string.Format(CultureInfo.CurrentCulture, Properties.Resources.FileNotValidXml, ex.Message));
                        ResetTimer();
                    }

                    //and refresh to it!
                    if (bNewStream)
                    {
                        ControlMain controlMain = GetMainControl();
                        controlMain.RefreshControls(ControlMain.ChangeReason.DragDrop, null, null, null, null, null);
                    }
                }
            }
        }

        private void treeView_DragEnter(object sender, DragEventArgs e)
        {
            log.Debug("treeView_DragEnter fired");
            e.Effect = DragDropEffects.Copy;
        }

        private void treeView_MouseDown(object sender, MouseEventArgs e)
        {
            treeView.SelectedNode = treeView.GetNodeAt(e.X, e.Y);
        }

        private void treeView_ItemDrag(object sender, ItemDragEventArgs e)
        {
            openDopeDragHandler.treeView_ItemDrag(sender, e, 
                this,
                GetMainControl(), CurrentDocument, 
                CurrentPart, OwnerDocument,
                _PictureContentControlsReplace);
        }

        private void treeView_AfterSelect(object sender, TreeViewEventArgs e)
        {
            RefreshProperties();
        }

        private void treeView_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (XmlTreeIsEditable 
                    && ((XmlNode)treeView.SelectedNode.Tag).NodeType == XmlNodeType.Text)
            {
                e.Node.TreeView.LabelEdit = true;
                e.Node.BeginEdit();
            }
        }

        private void treeView_AfterLabelEdit(object sender, NodeLabelEditEventArgs e)
        {
            e.Node.TreeView.LabelEdit = false;

            try
            {
                Office.CustomXMLNode cxn = Utilities.MxnFromTn(e.Node, CurrentPart, false /* fRemoveTextNode */);
                cxn.Text = e.Label;
            }
            catch (COMException cex)
            {
                ShowErrorMessage(string.Format(CultureInfo.CurrentUICulture, Properties.Resources.ErrorRenamingNode, cex.Message));
                e.CancelEdit = true;
                return;
            }

        }

        #endregion

        #region Add methods

        /// <summary>
        /// Add a node to the tree view.
        /// </summary>
        /// <param name="addPosition">An AddPosition value determining where the new node should be added.</param>
        private void AddNodeFromDialog(Forms.FormAddNode.AddPosition addPosition)
        {
            using (Forms.FormAddNode fan = new Forms.FormAddNode((XmlNode)treeView.SelectedNode.Tag, CurrentPart))
            {
                if (fan.ShowDialog() == DialogResult.OK)
                {
                    //we have to do this in three phases:
                    // 1) add it to our DOM
                    // 2) try to add it to Office
                    // 3) if that fails, roll back the change to Office
                    //
                    //this is because #2 might result in either another change (which needs an up-to-date DOM),
                    // or an error that needs it to be rolled back 

                    //befor we start, get parent and next sibling
                    XmlNode xnNew = fan.NodeToImport;
                    XmlNode xnParent = null;
                    XmlNode xnNextSibling = null;
                    Office.CustomXMLNode mxnParent = null;
                    Office.CustomXMLNode mxnNextSibling = null;
                    switch (addPosition)
                    {
                        case Forms.FormAddNode.AddPosition.AppendChild:
                            xnParent = (XmlNode)treeView.SelectedNode.Tag;
                            mxnParent = Utilities.MxnFromTn(treeView.SelectedNode, CurrentPart, false);
                            break;
                        case Forms.FormAddNode.AddPosition.InsertBefore:
                            xnParent = (XmlNode)treeView.SelectedNode.Parent.Tag;
                            mxnParent = Utilities.MxnFromTn(treeView.SelectedNode.Parent, CurrentPart, false);
                            xnNextSibling = (XmlNode)treeView.SelectedNode.Tag;
                            mxnNextSibling = Utilities.MxnFromTn(treeView.SelectedNode, CurrentPart, false);
                            break;
                        case Forms.FormAddNode.AddPosition.InsertAfter:
                            xnParent = (XmlNode)treeView.SelectedNode.Parent.Tag;
                            mxnParent = Utilities.MxnFromTn(treeView.SelectedNode.Parent, CurrentPart, false);
                            xnNextSibling = ((XmlNode)treeView.SelectedNode.Tag).NextSibling;
                            mxnNextSibling = Utilities.MxnFromTn(treeView.SelectedNode, CurrentPart, false).NextSibling;
                            break;
                    }

                    //#1 - add it to our local DOM
                    XmlNode xn = OwnerDocument.ImportNode(xnNew, true);
                    AddNodeFromDialogToDom(xnParent, ref xnNextSibling, ref xn);
                    AddNodeToTree(xnParent, xnNextSibling, xn);
                    
                    //#2 - try to add it to Office                    
                    try
                    {                        
                        if (xnNew.NodeType == XmlNodeType.Element)
                        {
                            switch (addPosition)
                            {
                                case Forms.FormAddNode.AddPosition.AppendChild:
                                    mxnParent.AppendChildSubtree(xnNew.OuterXml);
                                    break;
                                default:
                                    mxnParent.InsertSubtreeBefore(xnNew.OuterXml, mxnNextSibling);
                                    break;
                            }
                        }
                        else if (xnNew.NodeType == XmlNodeType.Text && mxnNextSibling == null && mxnParent.LastChild != null && mxnParent.LastChild.NodeType == Office.MsoCustomXMLNodeType.msoCustomXMLNodeText)
                        {
                            //we want to append a text node after a text node
                            //change the value of the existing one
                            mxnParent.LastChild.Text = mxnParent.LastChild.Text + xnNew.InnerText;
                        }
                        else if (xnNew.NodeType == XmlNodeType.Text && mxnNextSibling != null && mxnNextSibling.NodeType == Office.MsoCustomXMLNodeType.msoCustomXMLNodeText)
                        {
                            //we want to add a text node ahead of another text node
                            //change the value of the existing one
                            mxnNextSibling.Text = xnNew.InnerText + mxnNextSibling.Text;
                        }
                        else if (xnNew.NodeType == XmlNodeType.Text && mxnNextSibling != null && mxnNextSibling.PreviousSibling != null && mxnNextSibling.PreviousSibling.NodeType == Office.MsoCustomXMLNodeType.msoCustomXMLNodeText)
                        {
                            //we want to add a text node after another text node
                            //change the value of the existing one
                            mxnNextSibling.PreviousSibling.Text = xnNew.InnerText + xnNew.InnerText;
                        }
                        else if (xnParent.NodeType == XmlNodeType.Attribute)
                        {
                            //since we cannot get to the text node in the Office CustomXMLPart, we'll replace the attribute's text instead
                            mxnParent.Text = xnNew.InnerText;
                        }
                        else
                        {
                            CurrentPart.AddNode(mxnParent, xnNew.Name, xnNew.NamespaceURI, mxnNextSibling, Utilities.CxntFromXnt(xnNew.NodeType), xnNew.InnerText);
                        }
                    }
                    //#3 - if an error, roll back
                    catch (COMException cex)
                    {
                        //remove it from our DOM
                        DeleteNode(xn, xnParent);

                        ShowErrorMessage(string.Format(CultureInfo.CurrentCulture, Properties.Resources.ErrorAddingNode, cex.Message));
                        return;
                    }
                    catch (ArgumentException ex)
                    {
                        //remove it from our DOM
                        DeleteNode(xn, xnParent);

                        ShowErrorMessage(string.Format(CultureInfo.CurrentCulture, Properties.Resources.ErrorAddingNode, ex.Message));
                        return;
                    }

                    //refresh properties if needed
                    if (treeView.SelectedNode != null)
                    {
                        while (xn != null)
                            if (xnNew.ParentNode == (XmlNode)treeView.SelectedNode.Tag)
                            {
                                RefreshProperties();
                                break;
                            }
                            else
                            {
                                xn = xn.ParentNode;
                            }
                    }
                }
            }
        }

        /// <summary>
        /// Add a node to the XML tree.
        /// </summary>
        /// <param name="xnParent">An XmlNode specifying the parent of the new node.</param>
        /// <param name="xnNextSibling">An XmlNode specifying the next sibling of the new node.</param>
        /// <param name="xn">An XmlNode specifying the new node.</param>
        private void AddNodeFromDialogToDom(XmlNode xnParent, ref XmlNode xnNextSibling, ref XmlNode xn)
        {
            if (xn.NodeType == XmlNodeType.Attribute)
            {
                xn = xnParent.Attributes.Append((XmlAttribute)xn);
            }
            else if (xn.NodeType == XmlNodeType.Text && xnNextSibling == null && xnParent.LastChild != null && xnParent.LastChild.NodeType == XmlNodeType.Text)
            {
                //we want to append a text node after a text node
                //change the value of the existing one
                xnParent.LastChild.InnerText = xnParent.LastChild.InnerText + xn.InnerText;

                //delete it (it will get added back later as part of adding it to the treeview)
                try
                {
                    m_dicXnTn[xnParent.LastChild].Remove();
                    m_dicXnTn.Remove(xnParent.LastChild);
                }
                catch (KeyNotFoundException) { // Happens if we explicitly add a text node, having already added the element
                }
                //let it get added back
                xn = xnParent.LastChild;
            }
            else if (xn.NodeType == XmlNodeType.Text && xnNextSibling != null && xnNextSibling.NodeType == XmlNodeType.Text)
            {
                //we want to add a text node ahead of another text node
                //change the value of the existing one
                xnNextSibling.InnerText = xn.InnerText + xnNextSibling.InnerText;

                //delete it (it will get added back later as part of adding it to the treeview)
                m_dicXnTn[xnNextSibling].Remove();
                m_dicXnTn.Remove(xnNextSibling);

                //let it get added back
                xn = xnNextSibling;
                xnNextSibling = xnNextSibling.NextSibling;
            }
            else if (xn.NodeType == XmlNodeType.Text && xnNextSibling != null && xnNextSibling.PreviousSibling != null && xnNextSibling.PreviousSibling.NodeType == XmlNodeType.Text)
            {
                //we want to add a text node after another text node
                //change the value of the existing one
                xnNextSibling.PreviousSibling.InnerText = xnNextSibling.PreviousSibling.InnerText + xn.InnerText;

                //delete it
                m_dicXnTn[xnNextSibling.PreviousSibling].Remove();
                m_dicXnTn.Remove(xnNextSibling.PreviousSibling);

                //let it get added back
                xn = xnNextSibling.PreviousSibling;
            }
            else
            {
                xn = xnParent.InsertBefore(xn, xnNextSibling);
            }

            xn.OwnerDocument.Normalize();
        }

        #endregion
        
        #region Delete methods

        /// <summary>
        /// Delete the XML node corresponding to the selected node in the tree view.
        /// </summary>
        private void DeleteSelectedTreeNode()
        {
            //prep: get the xml node
            Office.CustomXMLNode mxnSelected = Utilities.MxnFromTn(treeView.SelectedNode, CurrentPart, false);
            Debug.Assert(mxnSelected != null, "ASSERT: null mxn", "This XPath didn't get a node: " + Utilities.XpathFromTn(treeView.SelectedNode, false, CurrentPart, null));

            //#1 - delete from our DOM
            XmlNode xn = (XmlNode)treeView.SelectedNode.Tag;
            XmlNode xnParent = xn.NodeType == XmlNodeType.Attribute ? ((XmlAttribute)xn).OwnerElement : xn.ParentNode;
            XmlNode xnNextSibling = xn.NextSibling;
            DeleteNode(xn, xnParent);

            //#2 - try to delete from Office                
            try
            {
                //special case, text node of attribute
                if (((XmlNode)treeView.SelectedNode.Tag).NodeType == XmlNodeType.Text && ((XmlNode)treeView.SelectedNode.Tag).ParentNode.NodeType == XmlNodeType.Attribute)
                {
                    //replace the text node
                    mxnSelected.Text = "";
                }
                else
                {
                    mxnSelected.Delete();
                }
            }
            //#3 - if error, roll back
            catch (COMException ex)
            {
                //add the node back
                AddNodeFromDialogToDom(xnParent, ref xnNextSibling, ref xn);
                AddNodeToTree(xnParent, xnNextSibling, xn);

                ShowErrorMessage(string.Format(CultureInfo.CurrentUICulture, Properties.Resources.ErrorDeletingNode, ex.Message));
            }
            catch (ArgumentException ex)
            {
                //add the node back
                AddNodeFromDialogToDom(xnParent, ref xnNextSibling, ref xn);
                AddNodeToTree(xnParent, xnNextSibling, xn);

                ShowErrorMessage(string.Format(CultureInfo.CurrentUICulture, Properties.Resources.ErrorDeletingNode, ex.Message));
            }

            RefreshProperties();
        }

        /// <summary>
        /// Delete the specified XML node from our XML tree.
        /// </summary>
        /// <param name="xn">The XmlNode to be deleted.</param>
        /// <param name="xnParent">The XmlNode that is the parent of the deleted node.</param>
        private void DeleteNode(XmlNode xn, XmlNode xnParent)
        {
            if (xn.NodeType == XmlNodeType.Attribute)
            {
                //remove the attribute
                if (m_dicXnTn.ContainsKey(xn))
                {
                    RemoveTreeNodeOnDelete(m_dicXnTn[xn]);
                    RemoveXmlNodeData(xn);
                }
                ((XmlNode)treeView.SelectedNode.Parent.Tag).Attributes.Remove((XmlAttribute)xn);                
            }
            else
            {
                //check if there are two text nodes to merge in the tree view now
                if (xn.PreviousSibling != null && xn.NextSibling != null && xn.PreviousSibling.NodeType == xn.NextSibling.NodeType && xn.PreviousSibling.NodeType == XmlNodeType.Text)
                {
                    //cache parent and next                        
                    XmlNode xnPreviousSibling = xn.PreviousSibling;

                    //clean up the lists
                    m_dicXnTn.Remove(xn);
                    m_dicXnTn.Remove(xn.PreviousSibling);
                    m_dicXnTn.Remove(xn.NextSibling);

                    //remove as asked
                    xn.ParentNode.RemoveChild(xn);

                    //normalize the DOM
                    xn.OwnerDocument.Normalize();

                    //clean up the tree
                    RemoveTreeNodeOnDelete(treeView.SelectedNode.PrevNode);
                    RemoveTreeNodeOnDelete(treeView.SelectedNode.NextNode);
                    RemoveTreeNodeOnDelete(treeView.SelectedNode);

                    //let it get added back
                    AddNodeToTree(xnParent, xnPreviousSibling.NextSibling, xnPreviousSibling);
                }
                else
                {
                    // just remove as asked
                    if(m_dicXnTn.ContainsKey(xn))
                    {
                        RemoveTreeNodeOnDelete(m_dicXnTn[xn]);
                        RemoveXmlNodeData(xn);
                    }
                    xn.ParentNode.RemoveChild(xn);                    
                }
            }
        }
        
        /// <summary>
        /// Remove a tree node from the tree.
        /// </summary>
        /// <param name="tn">The TreeNode to be removed.</param>
        private void RemoveTreeNodeOnDelete(TreeNode tn)
        {
            //cache the parent
            TreeNode tnParent;
            tnParent = tn.Parent;

            //remove the node
            RemoveTreeNodeData(tn);

            //fix the parent's icon
            if (tnParent != null && ((XmlNode)tnParent.Tag).NodeType != XmlNodeType.Attribute)
            {
                bool bIsLeafNode = IsLeafNode(tnParent);

                if (!bIsLeafNode)
                {
                    tnParent.ImageIndex = 0;
                    tnParent.SelectedImageIndex = 0;
                }
                else
                {
                    tnParent.ImageIndex = 2;
                    tnParent.SelectedImageIndex = 2;
                }
            }
        }

        /// <summary>
        /// Remove the tree node data corresponding to a deleted XML node.
        /// </summary>
        /// <param name="tn">The TreeNode to remove.</param>
        private void RemoveTreeNodeData(TreeNode tn)
        {
            //if the tree node is null, bail
            if (tn == null)
                return;

            //are there child nodes?
            //if so, call this recursively for those
            foreach (TreeNode tnChild in tn.Nodes)
                RemoveTreeNodeData(tnChild);

            //otherwise, just remove it
            tn.Remove();
        }

        /// <summary>
        /// Clean up our XML tree after a deletion.
        /// </summary>
        /// <param name="xn">An XmlNode to be deleted.</param>
        private void RemoveXmlNodeData(XmlNode xn)
        {
            //are there child nodes?
            //if so, call this recursively for those
            foreach (XmlNode xnChild in xn.ChildNodes)
                RemoveXmlNodeData(xnChild);
            if(xn.Attributes != null)
                foreach (XmlNode xnAttribute in xn.Attributes)
                    RemoveXmlNodeData(xnAttribute);

            //otherwise, just remove it
            m_dicXnTn.Remove(xn);
        }

        #endregion

        #region Context Menu methods/events

        private void doLeafSettings()
        {
            if (this.GetMainControl().modeControlEnabled == false)
            {
                // For a leaf node, bind and condition makes sense
                this.bindToolStripMenuItem.Visible = true;
                this.conditionToolStripMenuItem.Visible = true;
                this.repeatToolStripMenuItem.Visible = false;

                Word.ContentControl parentCC = CurrentDocument.Application.Selection.ParentContentControl;
                if (parentCC == null)
                {
                    // Can insert only, since there is nothing to map
                    bindMapToSelectedControlToolStripMenuItem.Visible = false;
                    bindInsertToolStripMenuItem.Visible = true;

                    conditionMapToSelectedControlToolStripMenuItem.Visible = false;
                    conditionInsertToolStripMenuItem.Visible = true;

                }
                else if (ContentControlOpenDoPEType.isRepeat(parentCC)
                 || ContentControlOpenDoPEType.isCondition(parentCC))
                {
                    // can insert only .. changing a repeat or condition to a bind
                    // will cause probs (at least if it has child controls)
                    bindMapToSelectedControlToolStripMenuItem.Visible = false;
                    bindInsertToolStripMenuItem.Visible = true;

                    conditionMapToSelectedControlToolStripMenuItem.Visible = false;
                    conditionInsertToolStripMenuItem.Visible = true;
                }
                else if (ContentControlOpenDoPEType.isBound(parentCC) )
                {
                    this.conditionToolStripMenuItem.Visible = false;

                    bindMapToSelectedControlToolStripMenuItem.Visible = true;
                    bindInsertToolStripMenuItem.Visible = false;
                }
                else 
                {
                    // anything else ..
                    bindMapToSelectedControlToolStripMenuItem.Visible = true;
                    bindInsertToolStripMenuItem.Visible = true;

                    conditionMapToSelectedControlToolStripMenuItem.Visible = false;
                    conditionInsertToolStripMenuItem.Visible = true;
                }
            }

        }

        private void contextMenuNode_Opening(object sender, CancelEventArgs e)
        {
            if (treeView.SelectedNode == null)
            {
                e.Cancel = true;
                return;
            }

            SetDefaultContextMenuState();

            // Show the mapping option if we've got a content control selected
            bool mapExisting = mapExistingControlMenuItemVisible();
            if (this.GetMainControl().modeControlEnabled)
            {
                mapToSelectedControlToolStripMenuItem.Visible = mapExisting;
            }
            else 
            {
                mapToSelectedControlToolStripMenuItem.Visible = false;
                // Want these instead
                bindMapToSelectedControlToolStripMenuItem.Visible = mapExisting;
                conditionMapToSelectedControlToolStripMenuItem.Visible = mapExisting;
                repeatMapToSelectedControlToolStripMenuItem.Visible = mapExisting;
            }

            SetDefaultMappingCheckboxState();

            

            switch (((XmlNode)treeView.SelectedNode.Tag).NodeType)
            {
                case XmlNodeType.Element:
                    if (IsLeafNode(treeView.SelectedNode)) {

                        doLeafSettings();
                    }
                    else
                    {
                        mapToSelectedControlToolStripMenuItem.Visible = false;
                        if (this.GetMainControl().modeControlEnabled == false)
                        {
                            this.bindToolStripMenuItem.Visible = false;
                            this.conditionToolStripMenuItem.Visible = conditionViaRightClick; // true,if functionality enabled.
                            this.repeatToolStripMenuItem.Visible = true; // makes the most sense

                            Word.ContentControl parentCC = CurrentDocument.Application.Selection.ParentContentControl;
                            if (parentCC == null)
                            {
                                // Can insert only, since there is nothing to map
                                conditionMapToSelectedControlToolStripMenuItem.Visible = false;
                                conditionInsertToolStripMenuItem.Visible = true;

                                repeatMapToSelectedControlToolStripMenuItem.Visible = false;
                                repeatInsertToolStripMenuItem.Visible = true;
                            }
                            else 
                            {
                                // can insert or map
                                conditionMapToSelectedControlToolStripMenuItem.Visible = true;
                                conditionInsertToolStripMenuItem.Visible = true;

                                repeatMapToSelectedControlToolStripMenuItem.Visible = true;
                                repeatInsertToolStripMenuItem.Visible = true;
                            }

                        }
                        insertToolStripMenuItem.Visible = false;
                        toolStripSeparator1.Visible = false;
                    }
                    if (treeView.SelectedNode == m_tnRoot)
                    {
                        deleteToolStripMenuItem.Visible = false;
                        aboveToolStripMenuItem.Visible = false;
                        belowToolStripMenuItem.Visible = false;
                    }
                    break;
                case XmlNodeType.ProcessingInstruction:
                case XmlNodeType.CDATA:
                case XmlNodeType.Comment:
                    insertToolStripMenuItem.Visible = false;
                    mapToSelectedControlToolStripMenuItem.Visible = false;
                    if (this.GetMainControl().modeControlEnabled==false)
                    {
                        bindMapToSelectedControlToolStripMenuItem.Visible = false;
                        conditionMapToSelectedControlToolStripMenuItem.Visible = false;
                        repeatMapToSelectedControlToolStripMenuItem.Visible = false;
                    }
                    toolStripSeparator1.Visible = false;
                    insideToolStripMenuItem.Visible = false;
                    break;
                case XmlNodeType.Text:
                    if (IsLeafNode(treeView.SelectedNode.Parent)) {
                        doLeafSettings();
                    } else {
                        insertToolStripMenuItem.Visible = false;
                        mapToSelectedControlToolStripMenuItem.Visible = false;
                        if (this.GetMainControl().modeControlEnabled == false)
                        {
                            bindMapToSelectedControlToolStripMenuItem.Visible = false;
                            conditionMapToSelectedControlToolStripMenuItem.Visible = false;
                            repeatMapToSelectedControlToolStripMenuItem.Visible = false;
                        }
                        toolStripSeparator1.Visible = false;
                    }
                    insideToolStripMenuItem.Visible = false;
                    break;
                case XmlNodeType.Attribute:

                    aboveToolStripMenuItem.Visible = false;
                    belowToolStripMenuItem.Visible = false;

                    doLeafSettings();

                    break;
            }

            // hide delete on built-in parts
            if (CurrentPart.BuiltIn && ((XmlNode)treeView.SelectedNode.Tag).NodeType != XmlNodeType.Text)
            {
                //cannot delete or add above/below non-leaf elements
                deleteToolStripMenuItem.Visible = false;
                aboveToolStripMenuItem.Visible = false;
                belowToolStripMenuItem.Visible = false;

                //we're going to hide everything for non-leaf nodes, so just don't show it
                if (!IsLeafNode(treeView.SelectedNode))
                {
                    e.Cancel = true;
                }
            }

            //sniff the default drag type and set up the menu correctly
            if (OwnerDocument.Schemas.Count > 0)
            {
                switch (Utilities.CheckNodeType((XmlNode)treeView.SelectedNode.Tag))
                {
                    case Utilities.MappingType.Date:
                        dateDefaultToolStripMenuItem.Visible = true;
                        toolStripSeparator.Visible = true;
                        break;
                    case Utilities.MappingType.DropDown:
                        dropDownListDefaultToolStripMenuItem.Visible = true;
                        toolStripSeparator.Visible = true;
                        break;
                    case Utilities.MappingType.Picture:
                        pictureDefaultToolStripMenuItem.Visible = true;
                        toolStripSeparator.Visible = true;
                        break;
                    default:
                        break;
                }
            }
            else
            {
                // OpenDoPE TODO .. if the node contains base64 encoded content ... 
            }
        }

        /// <summary>
        /// Set the default state of the checkbox for the context menu entry for the "map to selected node" item.
        /// </summary>
        private void SetDefaultMappingCheckboxState()
        {
            //log.Debug("CurrentDocument.Application.Selection.ContentControls.Count == " + CurrentDocument.Application.Selection.ContentControls.Count);
            //log.Debug("CurrentDocument.Application.Selection.ParentContentControl == " + CurrentDocument.Application.Selection.ParentContentControl);

            mapToSelectedControlToolStripMenuItem.Checked = false;
            if (CurrentDocument.Application.Selection.ParentContentControl != null // ie selection is inside a content control
                  && CurrentDocument.Application.Selection.ParentContentControl.XMLMapping.IsMapped == true
                  && CurrentPart.Id == CurrentDocument.Application.Selection.ParentContentControl.XMLMapping.CustomXMLPart.Id
                  && ((XmlNode)treeView.SelectedNode.Tag == Utilities.XnFromMxn(OwnerDocument, CurrentDocument.Application.Selection.ParentContentControl.XMLMapping.CustomXMLNode, null)
                  || ((XmlNode)treeView.SelectedNode.Tag).NodeType == XmlNodeType.Text
                       && treeView.SelectedNode.Parent != null
                       && (XmlNode)treeView.SelectedNode.Parent.Tag == Utilities.XnFromMxn(OwnerDocument, CurrentDocument.Application.Selection.ParentContentControl.XMLMapping.CustomXMLNode, null)))
            {
                mapToSelectedControlToolStripMenuItem.Checked = true;

            }
            else if (CurrentDocument.Application.Selection.ContentControls.Count == 1)  // selection includes one or both *ends* of a content control
            {
                object objOne = 1;
                if (CurrentDocument.Application.Selection.ContentControls.get_Item(ref objOne).XMLMapping.IsMapped == true
                      && CurrentPart.Id == CurrentDocument.Application.Selection.ContentControls.get_Item(ref objOne).XMLMapping.CustomXMLPart.Id
                      && ((XmlNode)treeView.SelectedNode.Tag == Utilities.XnFromMxn(OwnerDocument, CurrentDocument.Application.Selection.ContentControls.get_Item(ref objOne).XMLMapping.CustomXMLNode, null)
                      || ((XmlNode)treeView.SelectedNode.Tag).NodeType == XmlNodeType.Text
                           && treeView.SelectedNode.Parent != null
                           && (XmlNode)treeView.SelectedNode.Parent.Tag == Utilities.XnFromMxn(OwnerDocument, CurrentDocument.Application.Selection.ContentControls.get_Item(ref objOne).XMLMapping.CustomXMLNode, null)))
                {
                    mapToSelectedControlToolStripMenuItem.Checked = true;
                }
            }
        }

        /// <summary>
        /// Set the default state of the context menu, based on the selected TreeNode.
        /// </summary>
        private void SetDefaultContextMenuState()
        {
            //set default state            

            toolStripSeparator1.Visible = false;
            addToolStripMenuItem.Visible = false;
            aboveToolStripMenuItem.Visible = false;
            belowToolStripMenuItem.Visible = false;
            insideToolStripMenuItem.Visible = false;
            deleteToolStripMenuItem.Visible = false;

            dateDefaultToolStripMenuItem.Visible = false;
            dropDownListDefaultToolStripMenuItem.Visible = false;
            pictureDefaultToolStripMenuItem.Visible = false;
            textDefaultToolStripMenuItem.Visible = false;
            toolStripSeparator.Visible = false;

            if (this.BindControlTypeChoice)
            {
                insertToolStripMenuItem.Visible = true;
                // choose between text, date, drop down list, picture, combo box
            }
            else
            {
                insertToolStripMenuItem.Visible = false;
            }
            mapToSelectedControlToolStripMenuItem.Visible = true;

            if (this.GetMainControl().modeControlEnabled) // usual config
            {
                this.bindToolStripMenuItem.Visible = false;
                this.conditionToolStripMenuItem.Visible = false;
                this.repeatToolStripMenuItem.Visible = false;
            }
            else
            {
                this.bindToolStripMenuItem.Visible = false; // altered if its a text node
                this.conditionToolStripMenuItem.Visible = conditionViaRightClick; // true, if functionality enabled
                this.repeatToolStripMenuItem.Visible = true;

                // Each of those has a submenu
                bindMapToSelectedControlToolStripMenuItem.Visible = true;
                bindInsertToolStripMenuItem.Visible = true;

                conditionMapToSelectedControlToolStripMenuItem.Visible = true;
                conditionInsertToolStripMenuItem.Visible = true;

                repeatMapToSelectedControlToolStripMenuItem.Visible = true;
                repeatInsertToolStripMenuItem.Visible = true; 

            }

            if (XmlTreeIsEditable)
            {
                toolStripSeparator1.Visible = true;
                addToolStripMenuItem.Visible = true;
                aboveToolStripMenuItem.Visible = true;
                belowToolStripMenuItem.Visible = true;
                insideToolStripMenuItem.Visible = true;
                deleteToolStripMenuItem.Visible = true;
            }
        }


        private void insideToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddNodeFromDialog(Forms.FormAddNode.AddPosition.AppendChild);
        }

        private void belowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddNodeFromDialog(Forms.FormAddNode.AddPosition.InsertAfter);
        }

        private void aboveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddNodeFromDialog(Forms.FormAddNode.AddPosition.InsertBefore);
        }

        private void textToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                CreateMappedControl(Word.WdContentControlType.wdContentControlText);
            }
            catch (COMException cex)
            {
                ShowErrorMessage(string.Format(CultureInfo.CurrentUICulture, Properties.Resources.ErrorAddPlainText, cex.Message));
            }
        }

        #region other control types
        private void dateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                CreateMappedControl(Word.WdContentControlType.wdContentControlDate);
            }
            catch (COMException cex)
            {
                ShowErrorMessage(string.Format(CultureInfo.CurrentUICulture, Properties.Resources.ErrorAddDate, cex.Message));
            }
        }

        private void dropDownListToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                CreateMappedControl(Word.WdContentControlType.wdContentControlDropdownList);
            }
            catch (COMException cex)
            {
                ShowErrorMessage(string.Format(CultureInfo.CurrentUICulture, Properties.Resources.ErrorAddDropDown, cex.Message));
            }
        }

        private void pictureToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                CreateMappedControl(Word.WdContentControlType.wdContentControlPicture);
            }
            catch (COMException cex)
            {
                ShowErrorMessage(string.Format(CultureInfo.CurrentUICulture, Properties.Resources.ErrorAddPicture, cex.Message));
            }
        }

        private void comboBoxToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                CreateMappedControl(Word.WdContentControlType.wdContentControlComboBox);
            }
            catch (COMException cex)
            {
                ShowErrorMessage(string.Format(CultureInfo.CurrentUICulture, Properties.Resources.ErrorAddComboBox, cex.Message));
            }
        }

        #endregion

        private bool mapExistingControlMenuItemVisible()
        {
                if (CurrentDocument.Application.Selection.ContentControls.Count == 1) // ie selection includes one or both ends of a single content control
                {
                    return true;
                }
                else if (CurrentDocument.Application.Selection.ParentContentControl != null) // selection is strictly inside at least one content control
                {
                    return true;
                }
                else
                {
                    return false;
                }

        }

        private void bindInsertToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                CreateMappedControl(Word.WdContentControlType.wdContentControlText, OpenDopeType.Bind);
            }
            catch (COMException cex)
            {
                log.Error(cex);
                ShowErrorMessage(string.Format(CultureInfo.CurrentUICulture, Properties.Resources.ErrorAddPlainText, cex.Message));
            }
        }
        private void conditionInsertToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                CreateMappedControl(Word.WdContentControlType.wdContentControlText, OpenDopeType.Condition);
            }
            catch (COMException cex)
            {
                ShowErrorMessage(string.Format(CultureInfo.CurrentUICulture, Properties.Resources.ErrorAddPlainText, cex.Message));
            }
        }
        private void repeatInsertToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                CreateMappedControl(Word.WdContentControlType.wdContentControlText, OpenDopeType.Repeat);
            }
            catch (COMException cex)
            {
                ShowErrorMessage(string.Format(CultureInfo.CurrentUICulture, Properties.Resources.ErrorAddPlainText, cex.Message));
            }
        }

        private void bindMapToSelectedControlToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mapToSelectedControl(OpenDopeType.Bind);
        }

        private void conditionMapToSelectedControlToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mapToSelectedControl(OpenDopeType.Condition);
        }

        private void repeatMapToSelectedControlToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mapToSelectedControl(OpenDopeType.Repeat);
        }

        private void mapToSelectedControlToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mapToSelectedControl(OpenDopeType.Unspecified);
        }


        /// <summary>
        /// used when they right click then select "map to"
        /// </summary>
        /// <param name="odType"></param>
        private void mapToSelectedControl( OpenDopeType odType) {

            openDopeRightClickHandler.mapToSelectedControl(odType,
                this,
                GetMainControl(), CurrentDocument, 
                CurrentPart, 
                //OwnerDocument,
                _PictureContentControlsReplace);
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (ShowYesNoMessage(Properties.Resources.DeleteNodeMessage) == DialogResult.Yes)
            {
                DeleteSelectedTreeNode();
            }
        }

        public enum OpenDopeType { Unspecified, Bind, Condition, Repeat };


        /// <summary>
        /// Create a content control mapped to the selected XML node.
        /// </summary>
        /// <param name="CCType">A WdContentControlType value specifying the type of control to create.</param>
        private void CreateMappedControl(Word.WdContentControlType CCType)
        {
            CreateMappedControl(CCType, OpenDopeType.Unspecified);
        }
        /// <summary>
        /// Create a content control mapped to the selected XML node.
        /// </summary>
        /// <param name="CCType">A WdContentControlType value specifying the type of control to create.</param>
        private void CreateMappedControl(Word.WdContentControlType CCType, OpenDopeType odType)
        {
            openDopeCreateMappedControl.CreateMappedControl(CCType, odType,
                this,
                GetMainControl(), CurrentDocument,
                CurrentPart,
                //OwnerDocument,
                _PictureContentControlsReplace);
        }



        #endregion

        #region Helper functions

        /// <summary>
        /// Refresh the property window, based on the selected tree view node.
        /// </summary>
        private void RefreshProperties()
        {
            //no need to refresh if the property window isn't showing
            if((Options & cOptionsShowPropertyPage) == 0)
                return;

            ControlMain main = GetMainControl();

            if (treeView.SelectedNode != null)
            {
                main.RefreshProperties(Utilities.MxnFromTn(treeView.SelectedNode, CurrentPart, false));
            }
            else
            {
                main.RefreshProperties(null);
            }
        }

        /// <summary>
        /// Get the ControlMain control.
        /// </summary>
        /// <returns>The ControlMain control.</returns>
        private ControlMain GetMainControl()
        {
            Control c = this.Parent;
            ControlMain controlMain = null;
            while (true)
            {
                controlMain = c as ControlMain;
                if (controlMain != null)
                {
                    break;
                }
                else
                {
                    log.Debug(c.GetType().FullName);

                    c = c.Parent;
                }
            }
            return controlMain;
        }

        /// <summary>
        /// Determines whether the current node is a leaf node within the tree (i.e. whether it contains only text nodes).
        /// </summary>
        /// <param name="tn">A TreeNode specifying the node to evaluate.</param>
        /// <returns>True if the node is a leaf node, False otherwise.</returns>
        public static bool IsLeafNode(TreeNode tn)
        {
            bool bIsLeafNode = true;
            bool bFoundTextNode = false;
            XmlNode xnChildNode = ((XmlNode)tn.Tag).FirstChild;
            while (xnChildNode != null && bIsLeafNode == true)
            {
                if (xnChildNode.NodeType == XmlNodeType.Element)
                {
                    bIsLeafNode = false;
                }
                else if (xnChildNode.NodeType == XmlNodeType.Text)
                {
                    if (bFoundTextNode)
                    {
                        bIsLeafNode = false;
                    }
                    else
                    {
                        bFoundTextNode = true;
                    }
                }
                xnChildNode = xnChildNode.NextSibling;
            }
            return bIsLeafNode;
        }

        //private static bool HasXHTMLContent(TreeNode tn)
        //{
        //    String content = ((XmlNode)tn.Tag).InnerText;

        //    log.Info(content);

        //    return ContentDetection.IsXHTMLContent(content);
        //}

        public static string EscapeXHTML(string nodeContent)
        {
            nodeContent = nodeContent.Replace("'", "&apos;");
            nodeContent = nodeContent.Replace("\"", "&quot;");
            nodeContent = nodeContent.Replace("&", "&amp;");
            nodeContent = nodeContent.Replace("<", "&lt;");
            nodeContent = nodeContent.Replace(">", "&gt;");

            return nodeContent;
        }

        /// <summary>
        /// Get the visiblity of a node type in the tree, based on the current options.
        /// </summary>
        /// <param name="xn">The XmlNode to be added to the tree view.</param>
        /// <returns>True if that node type is visible, False otherwise.</returns>
        private static bool TnVisibiltyFetchFromGrf(XmlNode xn, int Options)
        {
            switch (xn.NodeType)
            {
                case XmlNodeType.Element:
                    return true;
                case XmlNodeType.Attribute:
                    if ((Options & cOptionsShowAttributes) != 0)
                        return true;
                    else
                        return false;
                case XmlNodeType.Comment:
                    if ((Options & cOptionsShowComments) != 0)
                        return true;
                    else
                        return false;
                case XmlNodeType.ProcessingInstruction:
                    if ((Options & cOptionsShowPI) != 0)
                        return true;
                    else
                        return false;
                case XmlNodeType.Text:
                    if ((Options & cOptionsShowText) != 0)
                        return true;
                    else
                        return false;
                default:
                    Debug.Fail("we got asked about an unknown node type");
                    return false;
            }
        }

        /// <summary>
        /// Add the appropriate icon to a tree node.
        /// </summary>
        /// <param name="xn">A XmlNode specifying the node's contents.</param>
        /// <param name="tn">A TreeNode specifying the tree node to update.</param>
        private static void AddIconToTreeNode(XmlNode xn, TreeNode tn)
        {
            //set up the icon
            if (xn.NodeType == XmlNodeType.Element || xn.NodeType == XmlNodeType.Document)
            {
                bool bIsLeafNode = IsLeafNode(tn);

                if (!bIsLeafNode)
                {
                    tn.ImageIndex = 0;
                    tn.SelectedImageIndex = 0;
                }
                else
                {
                    tn.ImageIndex = 2;
                    tn.SelectedImageIndex = 2;
                }
            }
            else if (xn.NodeType == XmlNodeType.Attribute)
            {
                tn.ImageIndex = 3;
                tn.SelectedImageIndex = 3;
            }
            else
            {
                tn.ImageIndex = 1;
                tn.SelectedImageIndex = 1;
            }
        }

        #endregion
    }

    /// <summary>
    /// Information needed while the tree view is being built.
    /// </summary>
    internal class TreeViewElements
    {
        private readonly Office.CustomXMLPart m_cxpTreeView;
        private readonly XmlDocument m_xdocTreeView;
        private readonly Office.CustomXMLNode m_mxnNodeToSelect;

        public TreeViewElements(Office.CustomXMLPart cxpTreeView, Office.CustomXMLNode mxnNodeToSelect)
        {
            m_cxpTreeView = cxpTreeView;
            m_xdocTreeView = new XmlDocument();
            m_mxnNodeToSelect = mxnNodeToSelect;
        }

        public XmlDocument TreeViewDOM
        {
            get
            {
                return m_xdocTreeView;
            }
        }

        public Office.CustomXMLPart TreeViewPart
        {
            get
            {
                return m_cxpTreeView;
            }
        }

        public Office.CustomXMLNode NodeToSelect
        {
            get
            {
                return m_mxnNodeToSelect;
            }
        }
    }

    /// <summary>
    /// This exception is thrown when the treeview refresh process should be immediately aboted.
    /// </summary>
    [Serializable()]
    public class QuitNowException : System.Exception
    {
        public QuitNowException() { }

        public QuitNowException(string message): base(message) { }

        public QuitNowException(string message, Exception innerException) : base(message, innerException) { }

        protected QuitNowException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}