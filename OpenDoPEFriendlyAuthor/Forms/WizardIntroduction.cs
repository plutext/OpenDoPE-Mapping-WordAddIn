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
