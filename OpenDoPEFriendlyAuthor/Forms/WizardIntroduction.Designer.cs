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
namespace XmlMappingTaskPane.Forms
{
    partial class WizardIntroduction
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WizardIntroduction));
            this.labelTitle = new System.Windows.Forms.Label();
            this.radioButtonCopyFile = new System.Windows.Forms.RadioButton();
            this.radioButtonTypeText = new System.Windows.Forms.RadioButton();
            this.labelRadioButtonHeader = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // labelTitle
            // 
            resources.ApplyResources(this.labelTitle, "labelTitle");
            this.labelTitle.Name = "labelTitle";
            // 
            // radioButtonCopyFile
            // 
            resources.ApplyResources(this.radioButtonCopyFile, "radioButtonCopyFile");
            this.radioButtonCopyFile.Checked = true;
            this.radioButtonCopyFile.Name = "radioButtonCopyFile";
            this.radioButtonCopyFile.TabStop = true;
            this.radioButtonCopyFile.UseVisualStyleBackColor = true;
            this.radioButtonCopyFile.MouseClick += new System.Windows.Forms.MouseEventHandler(this.radioButtonTypeText_MouseClick);
            // 
            // radioButtonTypeText
            // 
            resources.ApplyResources(this.radioButtonTypeText, "radioButtonTypeText");
            this.radioButtonTypeText.Name = "radioButtonTypeText";
            this.radioButtonTypeText.TabStop = true;
            this.radioButtonTypeText.UseVisualStyleBackColor = true;
            this.radioButtonTypeText.MouseClick += new System.Windows.Forms.MouseEventHandler(this.radioButtonTypeText_MouseClick);
            // 
            // labelRadioButtonHeader
            // 
            resources.ApplyResources(this.labelRadioButtonHeader, "labelRadioButtonHeader");
            this.labelRadioButtonHeader.Name = "labelRadioButtonHeader";
            // 
            // WizardIntroduction
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.labelRadioButtonHeader);
            this.Controls.Add(this.radioButtonTypeText);
            this.Controls.Add(this.radioButtonCopyFile);
            this.Controls.Add(this.labelTitle);
            this.Name = "WizardIntroduction";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelTitle;
        private System.Windows.Forms.RadioButton radioButtonCopyFile;
        private System.Windows.Forms.RadioButton radioButtonTypeText;
        private System.Windows.Forms.Label labelRadioButtonHeader;
    }
}
