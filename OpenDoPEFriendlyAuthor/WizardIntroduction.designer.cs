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
