//Copyright (c) Microsoft Corporation.  All rights reserved.
using System;
using System.Windows.Forms;

namespace XmlMappingTaskPane.Forms
{
    public partial class WizardFromString : UserControl
    {
        IButtonControl buttonOK = null;

        public WizardFromString()
        {
            InitializeComponent();            
        }

        internal string XmlString
        {
            get
            {
                return textBoxXml.Text;
            }
        }

        internal void SetFocus()
        {
            textBoxXml.Focus();
        }

        private void WizardFromString_VisibleChanged(object sender, EventArgs e)
        {
            if (this.Visible)
            {
                //turn off ENTER to proceed, so you can have multiple lines of XML
                buttonOK = ((FormAddPart)Parent).AcceptButton;
                ((FormAddPart)Parent).AcceptButton = null;
            }
            else if (buttonOK != null)
            {
                //restore it
                ((FormAddPart)Parent).AcceptButton = buttonOK;
            }
        }
    }
}
