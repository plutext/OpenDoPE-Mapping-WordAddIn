//Copyright (c) Microsoft Corporation.  All rights reserved.
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace XmlMappingTaskPane.Controls
{
    public partial class ControlProperties : Controls.ControlBase
    {
        public ControlProperties()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Refresh the property grid.
        /// </summary>
        /// <param name="cxn">A CustomXMLNode specifying the node whose properties we want to use.</param>
        internal void RefreshProperties(Office.CustomXMLNode cxn)
        {
            listViewProperties.Items.Clear();
            if (cxn != null)
            {
                //refresh properties
                listViewProperties.Items.Add(new ListViewItem(new string[] { Properties.Resources.Namespace, cxn.NamespaceURI }));
                listViewProperties.Items.Add(new ListViewItem(new string[] { Properties.Resources.Type, cxn.NodeType.ToString() }));
                listViewProperties.Items.Add(new ListViewItem(new string[] { Properties.Resources.XPath, cxn.XPath }));
                listViewProperties.Items.Add(new ListViewItem(new string[] { Properties.Resources.Prefixes, Utilities.GetPrefixMappingsMxn(cxn) }));
                listViewProperties.Items.Add(new ListViewItem(new string[] { Properties.Resources.XML, cxn.XML }));
            }
        }

        #region Events

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(listViewProperties.SelectedItems[0].SubItems[1].Text, true);
        }

        private void contextMenuStrip_Opening(object sender, CancelEventArgs e)
        {
            if (listViewProperties.SelectedItems.Count == 0)
            {
                e.Cancel = true;
            }
        }

        #endregion
    }
}
