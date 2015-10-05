namespace XmlMappingTaskPane.Forms
{
    partial class FormSelectRepeatedElement
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
            this.listElementNames = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.labelXPath = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // listElementNames
            // 
            this.listElementNames.FormattingEnabled = true;
            this.listElementNames.Location = new System.Drawing.Point(153, 41);
            this.listElementNames.Name = "listElementNames";
            this.listElementNames.Size = new System.Drawing.Size(196, 95);
            this.listElementNames.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 41);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(122, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Which element repeats?";
            // 
            // button1
            // 
            this.button1.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.button1.Location = new System.Drawing.Point(153, 152);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "OK";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // labelXPath
            // 
            this.labelXPath.AutoSize = true;
            this.labelXPath.Location = new System.Drawing.Point(12, 15);
            this.labelXPath.MinimumSize = new System.Drawing.Size(400, 0);
            this.labelXPath.Name = "labelXPath";
            this.labelXPath.Size = new System.Drawing.Size(400, 13);
            this.labelXPath.TabIndex = 3;
            this.labelXPath.Text = "xpath goes here";
            // 
            // FormSelectRepeatedElement
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(457, 191);
            this.Controls.Add(this.labelXPath);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.listElementNames);
            this.Name = "FormSelectRepeatedElement";
            this.Text = "FormSelectRepeatedElement";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.ListBox listElementNames;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        public System.Windows.Forms.Label labelXPath;
    }
}