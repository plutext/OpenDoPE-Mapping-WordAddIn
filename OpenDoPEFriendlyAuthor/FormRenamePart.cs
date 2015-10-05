//Copyright (c) Microsoft Corporation.  All rights reserved.
using System;
using System.Windows.Forms;
using XmlMappingTaskPane.Controls;

namespace XmlMappingTaskPane.Forms
{
    public partial class FormRenamePart : Form
    {
        private int m_intLocale;
        private string m_strCurrentValue;
        private string m_strCurrentNamespace;

        public FormRenamePart()
        {
            InitializeComponent();

            //move default focus onto the text
            textBoxName.Focus();
            textBoxName.SelectAll(); 
        }

        public int CurrentLocale
        {
            get
            {
                return m_intLocale;
            }
            set
            {
                m_intLocale = value;
            }
        }

        public string CurrentName
        {
            get
            {
                return m_strCurrentValue;
            }
            set
            {
                m_strCurrentValue = value;
                textBoxName.Text = value;
            }
        }

        public string CurrentNamespace
        {
            get
            {
                return m_strCurrentNamespace;
            }
            set
            {
                m_strCurrentNamespace = value;
            }
        }     

        private void buttonOK_Click(object sender, EventArgs e)
        {
            if (string.Equals(textBoxName.Text, m_strCurrentValue, StringComparison.CurrentCulture))
            {
                this.Close();
            }
            else if (SchemaLibrary.SetAlias(m_strCurrentNamespace, textBoxName.Text, m_intLocale))
            {
                DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                GenericMessageBox.Show(this, Properties.Resources.ChangePartNameErrorMessage, Properties.Resources.DialogTitle, MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0);
            }
        }
    }
}
