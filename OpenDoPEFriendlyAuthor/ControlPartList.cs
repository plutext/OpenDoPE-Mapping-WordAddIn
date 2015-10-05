//Copyright (c) Microsoft Corporation.  All rights reserved.
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace XmlMappingTaskPane.Controls
{
    public partial class ControlPartList : Controls.ControlBase
    {
        private int m_intLastSelected = -1; // the index of the last selection in the dropdown list
        private Office.CustomXMLNode m_mxnToSelect; // the node to be selected in the taskpane
        private IDictionary<string, int> m_dicCurrentIndices = new Dictionary<string, int>(); //key = ID property for each CustomXMLPart object; value = its position in the dropdown list

        public ControlPartList()
        {
            InitializeComponent();
        }

        #region Refresh methods

        /// <summary>
        /// Refresh the dropdown list.
        /// </summary>
        /// <param name="bRebuildList">True if the list should be rebuilt from the data in the document, False otherwise.</param>
        /// <param name="bSelectFirstPart">True if the first part in the dropdown list should be selected after the refresh.</param>
        /// <param name="bSelectNewestPart">True if the last part in the dropdown list should be selected after the refresh.</param>
        /// <param name="strIdOfPartToSelect">A string specifying the ID of the part to be selected after the refresh.</param>
        /// <param name="strPartToIgnore">A string specifying the ID of a part to be ignored as part of the refresh (because it is being deleted).</param>
        /// <param name="mxnNodeToSelect">A CustomXMLNode specifying the node to be selected after the refresh.</param>
        internal void RefreshPartList(bool bRebuildList, bool bSelectFirstPart, bool bSelectNewestPart, string strIdOfPartToSelect, string strPartToIgnore, Office.CustomXMLNode mxnNodeToSelect)
        {
            this.SuspendLayout();

            // store the location of the deleted item if we're losing one, since we need to reposition the selection afterwards
            int intPositionOfDeletedItem = -1;
            if (!string.IsNullOrEmpty(strPartToIgnore))
            {
                intPositionOfDeletedItem = m_dicCurrentIndices[strPartToIgnore];
            }

            // store the node to select
            m_mxnToSelect = mxnNodeToSelect;

            // only rebuild if we need to
            if (bRebuildList)
            {
                RebuildPartList(strPartToIgnore);
            }

            // get the right selection
            if (bSelectFirstPart)
            {
                comboBoxPartList.SelectedIndex = 0;
            }
            else if (bSelectNewestPart)
            {
                Debug.Assert(m_intLastSelected != comboBoxPartList.Items.Count - 2, "PERF: Double Refresh?", "Why are we refreshing on newest stream index?" + m_intLastSelected.ToString(CultureInfo.InvariantCulture) + " to " + comboBoxPartList.SelectedIndex.ToString(CultureInfo.InvariantCulture));

                // move to the newest item
                // this will implicitly force a tree refresh, because we're switching the index
                Debug.WriteLine("Tree refresh in RefreshPartList w/ selectNewestPart.");
                comboBoxPartList.SelectedIndex = comboBoxPartList.Items.Count - 2;
            }
            else if (!string.IsNullOrEmpty(strIdOfPartToSelect))
            {
                if (comboBoxPartList.SelectedIndex == m_dicCurrentIndices[strIdOfPartToSelect])
                {
                    ((ControlMain)Parent).SelectNodeFromTree(m_mxnToSelect);
                }
                else
                {
                    comboBoxPartList.SelectedIndex = m_dicCurrentIndices[strIdOfPartToSelect];

                    if (m_intLastSelected != m_dicCurrentIndices[strIdOfPartToSelect])
                    {
                        // refresh the treeview
                        Debug.WriteLine("Tree refresh in RefreshStreamSelect w/ specific stream");
                        ((ControlMain)Parent).RefreshTreeControl(m_mxnToSelect);
                    }
                }
            }
            else
            {
                //we want to keep on the same stream
                if (!string.IsNullOrEmpty(strPartToIgnore) && intPositionOfDeletedItem < m_intLastSelected) //deleting one above, so move the selection up one
                {
                    comboBoxPartList.SelectedIndex = m_intLastSelected - 1;
                }
                else if (m_intLastSelected == intPositionOfDeletedItem) //deleting the one we're on, so go to the top
                {
                    comboBoxPartList.SelectedIndex = 0;
                }
                else if (m_intLastSelected != comboBoxPartList.SelectedIndex) //deleting one below, so do nothing
                {
                    comboBoxPartList.SelectedIndex = m_intLastSelected;
                }
            }

            //set the current selection
            m_intLastSelected = comboBoxPartList.SelectedIndex;

            //clear
            m_mxnToSelect = null;

            this.ResumeLayout();
            this.PerformLayout();
        }

        /// <summary>
        /// Rebuild the contents of the dropdown list.
        /// </summary>
        /// <param name="strPartToIgnore">A string specifying the ID of a part to be ignored as part of the refresh (because it is being deleted).</param>
        internal void RebuildPartList(string strPartToIgnore)
        {
            IDictionary<string, int> dicNSCount = new Dictionary<string, int>();

            //clear the dropdown
            comboBoxPartList.Items.Clear();
            m_dicCurrentIndices.Clear();

            //repopulate the dropdown
            int index = 0;
            foreach (Office.CustomXMLPart currentPart in CurrentPartCollection)
            {
                //set up the dropdown entry
                string strCurrentNamespace = currentPart.NamespaceURI;

                //try the schema library to see if we can get an alias
                string alias = SchemaLibrary.GetAlias(strCurrentNamespace, Globals.ThisAddIn.Application.LanguageSettings.get_LanguageID(Microsoft.Office.Core.MsoAppLanguageID.msoLanguageIDUI));

                //if we're in the process of deleting this one, then skip it when rebuilding the list
                if (!string.IsNullOrEmpty(strPartToIgnore) && currentPart.Id == strPartToIgnore)
                    continue;

                //if there is more than one, update the stored index and populate (with the alias if we have one, otherwise use the namespace)
                if (dicNSCount.ContainsKey(strCurrentNamespace))
                {
                    //first, update the stored index
                    int intCurrentNSCount = ++dicNSCount[strCurrentNamespace];

                    //now, populate the dropdown
                    if (string.IsNullOrEmpty(strCurrentNamespace))
                    {
                        comboBoxPartList.Items.Add(string.Format(CultureInfo.CurrentCulture, Properties.Resources.NumberedItem, new object[] { Properties.Resources.NoNamespace, intCurrentNSCount }));
                    }
                    else if (!string.IsNullOrEmpty(alias))
                    {
                        comboBoxPartList.Items.Add(string.Format(CultureInfo.CurrentCulture, Properties.Resources.NumberedItem, new object[] { alias, intCurrentNSCount }));
                    }
                    else
                    {
                        comboBoxPartList.Items.Add(string.Format(CultureInfo.CurrentCulture, Properties.Resources.NumberedItem, new object[] { strCurrentNamespace, intCurrentNSCount }));
                    }
                }
                else
                {
                    //only one - populate the dropdown
                    if (string.IsNullOrEmpty(strCurrentNamespace))
                    {
                        comboBoxPartList.Items.Add(Properties.Resources.NoNamespace);
                    }
                    else if (!string.IsNullOrEmpty(alias))
                    {
                        comboBoxPartList.Items.Add(alias);
                    }
                    else
                    {
                        comboBoxPartList.Items.Add(strCurrentNamespace);
                    }

                    //add to the list
                    dicNSCount.Add(strCurrentNamespace, 1);
                }

                //add the entry to the hashtable
                m_dicCurrentIndices.Add(currentPart.Id, index);
                ++index;
            }

            //add the 'add new' entry last
            comboBoxPartList.Items.Add(Properties.Resources.AddNewDataSource);
        }

        #endregion

        #region Events

        private void comboBoxPartList_SelectedIndexChanged(object sender, EventArgs e)
        {
            //selected entry was changed
            //check if the selection didn't actually change
            if (comboBoxPartList.SelectedIndex == m_intLastSelected && m_dicCurrentIndices.ContainsKey(CurrentPart.Id) && CurrentPart.DocumentElement.OwnerDocument == CurrentPartCollection[m_dicCurrentIndices[CurrentPart.Id] + 1].DocumentElement.OwnerDocument)
            {
                return;
            }

            Debug.WriteLine("Switched selection in stream list.");

            //clear the tooltip
            toolTipNamespace.Active = false;
            toolTipNamespace.RemoveAll();

            //reset property page
            ((ControlMain)Parent).RefreshProperties(null);

            //check if we selected the last item in the list
            if (comboBoxPartList.SelectedIndex == comboBoxPartList.Items.Count - 1)
            {
                //then we selected add
                using (Forms.FormAddPart fap = new Forms.FormAddPart())
                {
                    //add a new stream from the XML retrieved from the Add New dialog
                    //otherwise, select the last selected item and populate with its xml
                    if (fap.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            //add the stream
                            object objMissing = Type.Missing;
                            CurrentPartCollection.Add(fap.XmlString, objMissing);

                            Debug.WriteLine("Dropdown refresh for manual stream addition");
                            RefreshPartList(false, false, true, string.Empty, string.Empty, null);
                        }
                        catch (COMException ex)
                        {
                            ShowErrorMessage(string.Format(CultureInfo.CurrentCulture, Properties.Resources.ErrorOnPartAdd, ex.Message));
                            if (comboBoxPartList.SelectedIndex != m_intLastSelected)
                            {
                                comboBoxPartList.SelectedIndex = m_intLastSelected;
                            }
                        }
                    }
                    else
                    {
                        //user cancel
                        if (comboBoxPartList.SelectedIndex != m_intLastSelected)
                        {
                            comboBoxPartList.SelectedIndex = m_intLastSelected;
                        }
                    }
                }
            }
            else
            {
                Debug.WriteLine("Tree refresh for SelectedIndexChanged from " + m_intLastSelected.ToString(CultureInfo.InvariantCulture) + " to " + comboBoxPartList.SelectedIndex.ToString(CultureInfo.InvariantCulture));

                //set the new selection
                m_intLastSelected = comboBoxPartList.SelectedIndex;

                //fire back at the event class to switch up the event handlers
                Debug.Assert(EventHandler != null, "null event handler");
                EventHandler.ChangeCurrentPart(CurrentPartCollection[m_intLastSelected + 1]);

                //refresh the tree
                ((ControlMain)Parent).RefreshTreeControl(m_mxnToSelect);
            }
        }

        private void comboBoxPartList_MouseHover(object sender, EventArgs e)
        {
            try
            {
                Debug.Assert(comboBoxPartList.SelectedIndex != -1, "SelectedIndex is null", "Why is there no SelectedIndex property set?");

                //if the current item isn't the add one, generate the tooltip
                if (comboBoxPartList.SelectedIndex != comboBoxPartList.Items.Count - 1)
                {
                    //create the tooltip
                    toolTipNamespace.Active = true;
                    toolTipNamespace.SetToolTip(comboBoxPartList, CurrentPartCollection[comboBoxPartList.SelectedIndex + 1].NamespaceURI);
                }
            }
            catch (COMException ex)
            {
                Debug.Fail(ex.Message);
            }
        }

        private void renameToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //allow them to rename the item
            using (Forms.FormRenamePart frmRename = new Forms.FormRenamePart())
            {
                frmRename.CurrentLocale = CurrentDocument.Application.LanguageSettings.get_LanguageID(Office.MsoAppLanguageID.msoLanguageIDUI);
                frmRename.CurrentNamespace = CurrentPart.NamespaceURI;
                frmRename.CurrentName = SchemaLibrary.GetAlias(frmRename.CurrentNamespace, frmRename.CurrentLocale);
                if (string.IsNullOrEmpty(frmRename.CurrentName))
                {
                    frmRename.CurrentName = frmRename.CurrentNamespace;
                }

                if (frmRename.ShowDialog() == DialogResult.OK)
                {
                    RefreshPartList(true, false, false, string.Empty, string.Empty, null);
                }
            }
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //prompt to delete the stream
            if (ShowYesNoMessage(Properties.Resources.DeletePartMessage) == DialogResult.Yes)
            {
                try
                {
                    //try to delete the stream
                    Office.CustomXMLPart partToDelete = CurrentPartCollection[m_intLastSelected + 1];
                    partToDelete.Delete();

                    //refresh the picker
                    RefreshPartList(true, true, false, string.Empty, null, null);

                    //set the new selection
                    m_intLastSelected = comboBoxPartList.SelectedIndex;

                    //fire back at the event class to switch up the event handlers as well
                    Debug.Assert(EventHandler != null, "null event handler");
                    EventHandler.ChangeCurrentPart(CurrentPartCollection[m_intLastSelected + 1]);
                }
                catch (COMException ex)
                {
                    ShowErrorMessage(string.Format(CultureInfo.CurrentUICulture, Properties.Resources.ErrorDeletePart, ex.Message));
                }
            }
        }

        private void contextMenuPart_Opening(object sender, CancelEventArgs e)
        {
            if (comboBoxPartList.SelectedItem.ToString().Contains(Properties.Resources.NoNamespace))
            {
                renameToolStripMenuItem.Visible = false;
            }
            else
            {
                renameToolStripMenuItem.Visible = true;
            }
        }

        #endregion
    }
}
