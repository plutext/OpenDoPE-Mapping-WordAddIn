//Copyright (c) Microsoft Corporation.  All rights reserved.
using System;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Win32;
using XmlMappingTaskPane.Controls;

namespace XmlMappingTaskPane.Forms
{
    public partial class FormOptions : Form
    {
        private int m_grfOptions;

        public FormOptions()
        {
            InitializeComponent();
        }

        internal int NewOptions
        {
            get
            {
                return m_grfOptions;
            }
        }

        private void FormOptions_Load(object sender, EventArgs e)
        {
            //set up the options
            try
            {
                m_grfOptions = (int)Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\XML Mapping Task Pane").GetValue("Options");
            }
            catch (NullReferenceException nrex)
            {
                Debug.Assert(false, "Regkey corruption", "either the user manually deleted the regkeys, or something bad happened." + Environment.NewLine + nrex.Message);
            }

            //options exist

            //set attributes
            if ((m_grfOptions & ControlTreeView.cOptionsShowAttributes) != 0)
                checkBoxAttributes.Checked = true;
            else
                checkBoxAttributes.Checked = false;

            //set comments
            if ((m_grfOptions & ControlTreeView.cOptionsShowComments) != 0)
                checkBoxComments.Checked = true;
            else
                checkBoxComments.Checked = false;

            //set PIs
            if ((m_grfOptions & ControlTreeView.cOptionsShowPI) != 0)
                checkBoxPI.Checked = true;
            else
                checkBoxPI.Checked = false;

            //set text
            if ((m_grfOptions & ControlTreeView.cOptionsShowText) != 0)
                checkBoxText.Checked = true;
            else
                checkBoxText.Checked = false;

            //set property page
            if ((m_grfOptions & ControlTreeView.cOptionsShowPropertyPage) != 0)
                checkBoxProperties.Checked = true;
            else
                checkBoxProperties.Checked = false;

            //set autoselect
            if ((m_grfOptions & ControlTreeView.cOptionsAutoSelectNode) != 0)
                checkBoxAutomaticallySelect.Checked = true;
            else
                checkBoxAutomaticallySelect.Checked = false;
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            //set options
            int grfOptions = 0;

            //set up the bitflag
            if (checkBoxAttributes.Checked == true)
                grfOptions = grfOptions | ControlTreeView.cOptionsShowAttributes;

            if (checkBoxComments.Checked == true)
                grfOptions = grfOptions | ControlTreeView.cOptionsShowComments;

            if (checkBoxPI.Checked == true)
                grfOptions = grfOptions | ControlTreeView.cOptionsShowPI;

            if (checkBoxText.Checked == true)
                grfOptions = grfOptions | ControlTreeView.cOptionsShowText;

            if (checkBoxProperties.Checked == true)
                grfOptions = grfOptions | ControlTreeView.cOptionsShowPropertyPage;

            if (checkBoxAutomaticallySelect.Checked == true)
                grfOptions = grfOptions | ControlTreeView.cOptionsAutoSelectNode;

            //persist the options
            if (grfOptions != m_grfOptions)
            {
                m_grfOptions = grfOptions;

                try
                {
                    if (Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\XML Mapping Task Pane") == null)
                        Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Office\XML Mapping Task Pane", RegistryKeyPermissionCheck.ReadWriteSubTree);

                    using (RegistryKey rk = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\XML Mapping Task Pane", true))
                        rk.SetValue("Options", m_grfOptions);
                }
                catch (System.Security.SecurityException)
                {                    
                    ControlBase.ShowErrorMessage(this, Properties.Resources.ErrorSaveSettings);
                }

                //set the result to OK
                DialogResult = DialogResult.OK;
            }

            //close the form
            this.Close();
        }
    }
}
