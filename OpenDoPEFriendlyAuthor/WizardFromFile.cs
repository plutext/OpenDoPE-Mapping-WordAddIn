//Copyright (c) Microsoft Corporation.  All rights reserved.
using System;
using System.Windows.Forms;

namespace XmlMappingTaskPane.Forms
{
    public partial class WizardFromFile : UserControl
    {
        public WizardFromFile()
        {
            InitializeComponent();
        }

        internal string FilePath
        {
            get
            {
                return textBoxFilePath.Text;
            }
        }

        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            //if the dialog is canceled, just exit
            if (openFileDialog.ShowDialog() != DialogResult.OK)
                return;

            //otherwise, insert the file name
            textBoxFilePath.Text = openFileDialog.FileName; 
        }
    }
}
