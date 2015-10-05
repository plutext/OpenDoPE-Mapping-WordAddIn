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
using System.Globalization;
using System.Windows.Forms;
using System.Xml;
using System.Xml.XPath;
using XmlMappingTaskPane.Controls;
using Office = Microsoft.Office.Core;

namespace XmlMappingTaskPane.Forms
{
    public partial class FormAddNode : Form
    {
        internal enum AddPosition { InsertBefore, InsertAfter, AppendChild };

        private XmlNode m_xn;
        private Office.CustomXMLPart m_cxp;

        private XmlDocument m_xdoc = new XmlDocument();
        private XmlNode m_xnToAdd;

        private bool m_fCancelClose;

        public FormAddNode(IXPathNavigable node, Office.CustomXMLPart part)
        {
            InitializeComponent();

            //set up state
            m_xn = node as XmlNode;
            m_cxp = part;
        }

        internal XmlNode NodeToImport
        {
            get
            {
                return m_xnToAdd;
            }
        }

        private void FormAddNode_Load(object sender, EventArgs e)
        {
            //based on the node and position we're sent, filter options appropriately
            switch (m_xn.NodeType)
            {
                case XmlNodeType.Element:
                    if (m_cxp.BuiltIn)
                    {
                        //allow text only
                        comboBoxType.Items.Clear();
                        comboBoxType.Items.Add(Properties.Resources.Text);
                    }
                    break;
                case XmlNodeType.Comment:
                case XmlNodeType.ProcessingInstruction:
                case XmlNodeType.CDATA:
                    break;
                case XmlNodeType.Attribute:
                    //allow text only
                    comboBoxType.Items.Clear();
                    comboBoxType.Items.Add(Properties.Resources.Text);
                    break;
                case XmlNodeType.Text:
                    if (m_xn.ParentNode.NodeType == XmlNodeType.Attribute || m_cxp.BuiltIn)
                    {
                        //this is an attribute's text
                        //allow text only
                        comboBoxType.Items.Clear();
                        comboBoxType.Items.Add(Properties.Resources.Text);
                    }
                    break;
                default:
                    Debug.Fail("why do we not handle this node type?");
                    break;
            }

            //preselect
            comboBoxType.SelectedIndex = 0;

            //set up the namespace list
            foreach (Office.CustomXMLPrefixMapping cxpm in m_cxp.NamespaceManager)
            {
                comboBoxNamespace.Items.Add(cxpm.NamespaceURI);
            }
            comboBoxNamespace.SelectedText = m_xn.NamespaceURI;

            SetControlStates();
        }

        /// <summary>
        /// Set the state of each control, based on the user's current selection of node type.
        /// </summary>
        private void SetControlStates()
        {
            //based on new selection, turn on/off controls
            if (comboBoxType.Text == Properties.Resources.Text || comboBoxType.Text == Properties.Resources.Comment || comboBoxType.Text == Properties.Resources.CDATA)
            {
                textBoxName.Clear();
                textBoxName.Enabled = false;
                comboBoxNamespace.SelectedText = string.Empty;
                comboBoxNamespace.Enabled = false;
                textBoxValue.Enabled = true;
            }
            else if (comboBoxType.Text == Properties.Resources.Element || comboBoxType.Text == Properties.Resources.Attribute)
            {
                textBoxName.Enabled = true;
                comboBoxNamespace.Enabled = true;
                textBoxValue.Enabled = true;
            }
            else if (comboBoxType.Text == Properties.Resources.ProcessingInstruction)
            {
                comboBoxNamespace.Text = string.Empty;
                comboBoxNamespace.Enabled = false;
            }
            else if (comboBoxType.Text == Properties.Resources.CompleteXMLDocument)
            {

            LShowDialog:

                //show open dialog
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //if they picked a file, then validate it, and add the filename to the value
                    try
                    {
                        m_xdoc.Load(openFileDialog.FileName);
                    }
                    catch (XmlException xex)
                    {
                        //if not valid, back to the dialog
                        ControlBase.ShowErrorMessage(this, string.Format(CultureInfo.CurrentCulture, Properties.Resources.FileNotValidXml, xex.Message));
                        goto LShowDialog;
                    }

                    //hide everything
                    textBoxName.Text = string.Empty;
                    textBoxName.Enabled = false;
                    comboBoxNamespace.Text = string.Empty;
                    comboBoxNamespace.Enabled = false;
                    textBoxValue.Text = openFileDialog.FileName;
                    textBoxValue.Enabled = false;
                }
                else
                {
                    //if cancel, go back to Element
                    comboBoxType.SelectedIndex = 0;
                }
            }
            else
            {
                Debug.Fail("why can't we handle this node type?");
            }
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            //check that input is valid
            if ((comboBoxType.Text == Properties.Resources.Element || comboBoxType.Text == Properties.Resources.Attribute) && string.IsNullOrEmpty(textBoxName.Text))
            {
                ControlBase.ShowErrorMessage(this, Properties.Resources.ErrorNoName);
                m_fCancelClose = true;
                return;
            }

            //add node(s)
            if (comboBoxType.Text == Properties.Resources.CompleteXMLDocument)
            {
                //build Xn
                m_xnToAdd = m_xdoc.DocumentElement;
            }
            else if (comboBoxType.Text == Properties.Resources.Comment)
            {
                m_xnToAdd = m_xdoc.CreateComment(textBoxValue.Text);
            }
            else if (comboBoxType.Text == Properties.Resources.CDATA)
            {
                m_xnToAdd = m_xdoc.CreateCDataSection(textBoxValue.Text);
            }
            else if (comboBoxType.Text == Properties.Resources.ProcessingInstruction)
            {
                m_xnToAdd = m_xdoc.CreateProcessingInstruction(textBoxName.Text, textBoxValue.Text);
            }
            else if (comboBoxType.Text == Properties.Resources.Text)
            {
                m_xnToAdd = m_xdoc.CreateTextNode(textBoxValue.Text);
            }
            else if (comboBoxType.Text == Properties.Resources.Attribute)
            {
                //add attr
                if (string.IsNullOrEmpty(comboBoxNamespace.Text))
                {
                    m_xnToAdd = m_xdoc.CreateAttribute(textBoxName.Text);
                }
                else if (!string.IsNullOrEmpty(m_cxp.NamespaceManager.LookupPrefix(comboBoxNamespace.Text)))
                {
                    m_xnToAdd = m_xdoc.CreateAttribute(m_cxp.NamespaceManager.LookupPrefix(comboBoxNamespace.Text), textBoxName.Text, comboBoxNamespace.Text);
                }
                else
                {
                    //find a new nsX value
                    int i = 0;
                    while (!string.IsNullOrEmpty(m_cxp.NamespaceManager.LookupNamespace("ns" + i)))
                        i++;

                    m_xnToAdd = m_xdoc.CreateAttribute("ns" + i, textBoxName.Text, comboBoxNamespace.Text);
                }

                m_xnToAdd.InnerText = textBoxValue.Text;
            }
            else if (comboBoxType.Text == Properties.Resources.Element)
            {
                //add element
                if (string.IsNullOrEmpty(comboBoxNamespace.Text))
                {
                    m_xnToAdd = m_xdoc.CreateElement(textBoxName.Text);
                }
                else if (!string.IsNullOrEmpty(m_cxp.NamespaceManager.LookupPrefix(comboBoxNamespace.Text)))
                {
                    m_xnToAdd = m_xdoc.CreateElement(m_cxp.NamespaceManager.LookupPrefix(comboBoxNamespace.Text), textBoxName.Text, comboBoxNamespace.Text);
                }
                else
                {
                    //find a new nsX value
                    int i = 0;
                    while (!string.IsNullOrEmpty(m_cxp.NamespaceManager.LookupNamespace("ns" + i)))
                        i++;

                    m_xnToAdd = m_xdoc.CreateElement("ns" + i, textBoxName.Text, comboBoxNamespace.Text);
                }
                m_xnToAdd.InnerText = textBoxValue.Text;
            }
            else
            {
                Debug.Fail("why can't we handle this node type?");
            }

            //hide, but don't close, this form
            DialogResult = DialogResult.OK;
            this.Hide();
        }

        private void FormAddNode_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (m_fCancelClose)
            {
                m_fCancelClose = false;
                e.Cancel = true;
            }
        }

        private void comboBoxType_SelectionChangeCommitted(object sender, EventArgs e)
        {
            SetControlStates();
        }

    }
}
