namespace XmlMappingTaskPane.Forms
{
    partial class FormOptions
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormOptions));
            this.checkBoxProperties = new System.Windows.Forms.CheckBox();
            this.checkBoxAutomaticallySelect = new System.Windows.Forms.CheckBox();
            this.groupBox = new System.Windows.Forms.GroupBox();
            this.checkBoxComments = new System.Windows.Forms.CheckBox();
            this.checkBoxPI = new System.Windows.Forms.CheckBox();
            this.checkBoxText = new System.Windows.Forms.CheckBox();
            this.checkBoxAttributes = new System.Windows.Forms.CheckBox();
            this.buttonOK = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.groupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // checkBoxProperties
            // 
            resources.ApplyResources(this.checkBoxProperties, "checkBoxProperties");
            this.checkBoxProperties.Name = "checkBoxProperties";
            this.checkBoxProperties.UseVisualStyleBackColor = true;
            // 
            // checkBoxAutomaticallySelect
            // 
            resources.ApplyResources(this.checkBoxAutomaticallySelect, "checkBoxAutomaticallySelect");
            this.checkBoxAutomaticallySelect.Name = "checkBoxAutomaticallySelect";
            this.checkBoxAutomaticallySelect.UseVisualStyleBackColor = true;
            // 
            // groupBox
            // 
            this.groupBox.Controls.Add(this.checkBoxComments);
            this.groupBox.Controls.Add(this.checkBoxPI);
            this.groupBox.Controls.Add(this.checkBoxText);
            this.groupBox.Controls.Add(this.checkBoxAttributes);
            resources.ApplyResources(this.groupBox, "groupBox");
            this.groupBox.Name = "groupBox";
            this.groupBox.TabStop = false;
            // 
            // checkBoxComments
            // 
            resources.ApplyResources(this.checkBoxComments, "checkBoxComments");
            this.checkBoxComments.Name = "checkBoxComments";
            this.checkBoxComments.UseVisualStyleBackColor = true;
            // 
            // checkBoxPI
            // 
            resources.ApplyResources(this.checkBoxPI, "checkBoxPI");
            this.checkBoxPI.Name = "checkBoxPI";
            this.checkBoxPI.UseVisualStyleBackColor = true;
            // 
            // checkBoxText
            // 
            resources.ApplyResources(this.checkBoxText, "checkBoxText");
            this.checkBoxText.Name = "checkBoxText";
            this.checkBoxText.UseVisualStyleBackColor = true;
            // 
            // checkBoxAttributes
            // 
            resources.ApplyResources(this.checkBoxAttributes, "checkBoxAttributes");
            this.checkBoxAttributes.Name = "checkBoxAttributes";
            this.checkBoxAttributes.UseVisualStyleBackColor = true;
            // 
            // buttonOK
            // 
            resources.ApplyResources(this.buttonOK, "buttonOK");
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            resources.ApplyResources(this.buttonCancel, "buttonCancel");
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            // 
            // FormOptions
            // 
            this.AcceptButton = this.buttonOK;
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.buttonCancel;
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.groupBox);
            this.Controls.Add(this.checkBoxAutomaticallySelect);
            this.Controls.Add(this.checkBoxProperties);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.HelpButton = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormOptions";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Load += new System.EventHandler(this.FormOptions_Load);
            this.groupBox.ResumeLayout(false);
            this.groupBox.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox checkBoxProperties;
        private System.Windows.Forms.CheckBox checkBoxAutomaticallySelect;
        private System.Windows.Forms.GroupBox groupBox;
        private System.Windows.Forms.CheckBox checkBoxComments;
        private System.Windows.Forms.CheckBox checkBoxPI;
        private System.Windows.Forms.CheckBox checkBoxText;
        private System.Windows.Forms.CheckBox checkBoxAttributes;
        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.Button buttonCancel;
    }
}