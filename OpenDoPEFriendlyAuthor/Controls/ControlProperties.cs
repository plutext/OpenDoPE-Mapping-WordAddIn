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
                if (show("NodeProperties.Show.Namespace"))
                    listViewProperties.Items.Add(new ListViewItem(new string[] { Properties.Resources.Namespace, cxn.NamespaceURI }));
                if (show("NodeProperties.Show.Type"))
                    listViewProperties.Items.Add(new ListViewItem(new string[] { Properties.Resources.Type, cxn.NodeType.ToString() }));
                if (show("NodeProperties.Show.XPath"))
                    listViewProperties.Items.Add(new ListViewItem(new string[] { Properties.Resources.XPath, cxn.XPath }));
                if (show("NodeProperties.Show.Prefixes"))
                    listViewProperties.Items.Add(new ListViewItem(new string[] { Properties.Resources.Prefixes, Utilities.GetPrefixMappingsMxn(cxn) }));
                if (show("NodeProperties.Show.XML"))
                    listViewProperties.Items.Add(new ListViewItem(new string[] { Properties.Resources.XML, cxn.XML }));
            }
        }

        private bool show(string property)
        {
            string val = System.Configuration.ConfigurationManager.AppSettings[property];
            if (String.IsNullOrWhiteSpace(val))
            {
                return false;
            }

            return val.ToLower().Equals("true");
        }

        internal void clear()
        {
            listViewProperties.Items.Clear();
        }

        internal void XPathWarning(string xpath)
        {
            listViewProperties.Items.Clear();
            {
                if (show("NodeProperties.Show.XPath"))
                {
                    ListViewItem lvi = new ListViewItem(new string[] { Properties.Resources.XPath, xpath });
                    lvi.BackColor = Color.Red;
                    listViewProperties.Items.Add(lvi);
                }
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
