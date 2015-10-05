//Copyright (c) Microsoft Corporation.  All rights reserved.
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
