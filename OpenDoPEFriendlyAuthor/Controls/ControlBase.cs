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
using System.Globalization;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace XmlMappingTaskPane.Controls
{
    public partial class ControlBase : UserControl
    {
        private DocumentEvents m_docEvents;

        public ControlBase()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Get the Document object associated with this control. Read-only.
        /// </summary>
        internal Word.Document CurrentDocument
        {
            get
            {
                return m_docEvents.Document;
            }
        }

        /// <summary>
        /// Get the CustomXMLParts object associated with the current document. Read-only.
        /// </summary>
        internal Office.CustomXMLParts CurrentPartCollection
        {
            get
            {
                return m_docEvents.PartCollection;
            }
        }

        /// <summary>
        /// Get the CustomXMLPart object associated with the currently selected part. Read-only.
        /// </summary>
        internal Office.CustomXMLPart CurrentPart
        {
            get
            {
                return m_docEvents.Part;
            }
        }

        /// <summary>
        /// Get/set the DocumentEvents object containing the document-level events for the current document.
        /// </summary>
        internal DocumentEvents EventHandler
        {
            get
            {
                return m_docEvents;
            }
            set
            {
                m_docEvents = value;
            }
        }

        #region Dialog box methods

        /// <summary>
        /// Show a standard error dialog box.
        /// </summary>
        /// <param name="Window">An IWin32Window object specifying the parent window for the dialog box.</param>
        /// <param name="Message">A string specifying the text to be shown in the dialog box .</param>
        internal static void ShowErrorMessage(IWin32Window Window, string Message)
        {
            GenericMessageBox.Show(Window, Message, Properties.Resources.DialogTitle, MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0);
        }

        /// <summary>
        /// Show a standard error dialog box, parented to the current control.
        /// </summary>
        /// <param name="Message">A string specifying the text to be shown in the dialog box .</param>
        internal void ShowErrorMessage(string Message)
        {
            GenericMessageBox.Show(this, Message, Properties.Resources.DialogTitle, MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0);
        }

        /// <summary>
        /// Show a standard "yes or no" dialog box, parented to the current control.
        /// </summary>
        /// <param name="Message">A string specifying the text to be shown in the dialog box .</param>
        /// <returns>A DialogResult specifying the button selected by the user.</returns>
        internal DialogResult ShowYesNoMessage(string Message)
        {
            return GenericMessageBox.Show(this, Message, Properties.Resources.DialogTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0);
        }

        #endregion
    }

    /// <summary>
    /// Specifies a message box that flips based on the reading order of the UI.
    /// </summary>
    public static class GenericMessageBox
    {
        public static DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton, MessageBoxOptions options)
        {
            if (IsRightToLeft(owner))
            {
                options |= MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign;
            }

            return MessageBox.Show(text, caption, buttons, icon, defaultButton, options);
        }
        
        private static bool IsRightToLeft(IWin32Window owner)
        {
            Control control = owner as Control;

            while (control != null)
            {
                if (control.RightToLeft != RightToLeft.Inherit)
                    return control.RightToLeft == RightToLeft.Yes;
                else
                    control = control.Parent;
            }

            // If no parent control is available, ask the CurrentUICulture
            // if we are running under right-to-left.
            return CultureInfo.CurrentUICulture.TextInfo.IsRightToLeft;
        }
    }
}
