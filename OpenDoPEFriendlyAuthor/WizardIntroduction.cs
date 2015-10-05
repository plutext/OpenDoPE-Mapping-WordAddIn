//Copyright (c) Microsoft Corporation.  All rights reserved.
using System;
using System.Windows.Forms;

namespace XmlMappingTaskPane.Forms
{
    public partial class WizardIntroduction : UserControl
    {
        internal enum RadioSelection { FromFile, FromString };

        private DateTime dateTimeLastClick = DateTime.MinValue;
        private object objLastControl = null;

        public WizardIntroduction()
        {
            InitializeComponent();
        }

        internal RadioSelection UserChoice
        {
            get
            {
                if (radioButtonCopyFile.Checked)
                    return RadioSelection.FromFile;
                else
                    return RadioSelection.FromString;
            }
        }

        private void radioButtonTypeText_MouseClick(object sender, MouseEventArgs e)
        {
            CheckForDoubleClick(sender);
        }

        private void radioButtonCopyFile_MouseClick(object sender, EventArgs e)
        {
            CheckForDoubleClick(sender);
        }

        private void CheckForDoubleClick(object sender)
        {
            //get current time
            DateTime dateTimeClick = DateTime.Now;
            System.Diagnostics.Debug.WriteLine(dateTimeClick);

            //check - was it within the double-click time on the system
            if (objLastControl == sender)
            {
                TimeSpan ts = dateTimeClick - dateTimeLastClick;
                System.Diagnostics.Debug.WriteLine(ts);
                if (ts.TotalMilliseconds <= (double)SystemInformation.DoubleClickTime)
                {
                    //yes, go to next page
                    ((FormAddPart)Parent).AcceptButton.PerformClick();
                }
            }

            //no, capture this one to compare to the next
            dateTimeLastClick = dateTimeClick;
            objLastControl = sender;
        }
    }
}
