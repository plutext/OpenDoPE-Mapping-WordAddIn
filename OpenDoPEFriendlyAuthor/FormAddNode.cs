//Copyright (c) Microsoft Corporation.  All rights reserved.
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
