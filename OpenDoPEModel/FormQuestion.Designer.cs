/*
 * (c) Copyright Plutext Pty Ltd, 2012
 * 
 * All rights reserved.
 * 
 * This source code is the proprietary information of Plutext
 * Pty Ltd, and must be kept confidential.
 * 
 * You may use, modify and distribute this source code only
 * as provided in your license agreement with Plutext.
 * 
 * If you do not have a license agreement with Plutext:
 * 
 * (i) you must return all copies of this source code to Plutext, 
 * or destroy it.  
 * 
 * (ii) under no circumstances may you use, modify or distribute 
 * this source code.
 * 
 */
namespace OpenDoPEModel
{
    partial class FormQuestion
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
            this.groupBoxQ = new System.Windows.Forms.GroupBox();
            this.textBoxQText = new System.Windows.Forms.TextBox();
            this.textBoxQID = new System.Windows.Forms.TextBox();
            this.labelQText = new System.Windows.Forms.Label();
            this.labelQID = new System.Windows.Forms.Label();
            this.groupBoxA = new System.Windows.Forms.GroupBox();
            this.radioButtonMCNo = new System.Windows.Forms.RadioButton();
            this.radioButtonMCYes = new System.Windows.Forms.RadioButton();
            this.labelAType = new System.Windows.Forms.Label();
            this.buttonNext = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.groupBoxXPath = new System.Windows.Forms.GroupBox();
            this.textBoxXPath = new System.Windows.Forms.TextBox();
            this.groupBoxQ.SuspendLayout();
            this.groupBoxA.SuspendLayout();
            this.groupBoxXPath.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBoxQ
            // 
            this.groupBoxQ.Controls.Add(this.textBoxQText);
            this.groupBoxQ.Controls.Add(this.textBoxQID);
            this.groupBoxQ.Controls.Add(this.labelQText);
            this.groupBoxQ.Controls.Add(this.labelQID);
            this.groupBoxQ.Location = new System.Drawing.Point(29, 78);
            this.groupBoxQ.Name = "groupBoxQ";
            this.groupBoxQ.Size = new System.Drawing.Size(367, 93);
            this.groupBoxQ.TabIndex = 0;
            this.groupBoxQ.TabStop = false;
            this.groupBoxQ.Text = "The Question";
            // 
            // textBoxQText
            // 
            this.textBoxQText.Location = new System.Drawing.Point(51, 53);
            this.textBoxQText.Name = "textBoxQText";
            this.textBoxQText.Size = new System.Drawing.Size(302, 20);
            this.textBoxQText.TabIndex = 3;
            this.textBoxQText.Text = "(type your question here)";
            // 
            // textBoxQID
            // 
            this.textBoxQID.Location = new System.Drawing.Point(51, 20);
            this.textBoxQID.Name = "textBoxQID";
            this.textBoxQID.Size = new System.Drawing.Size(100, 20);
            this.textBoxQID.TabIndex = 2;
            // 
            // labelQText
            // 
            this.labelQText.AutoSize = true;
            this.labelQText.Location = new System.Drawing.Point(17, 56);
            this.labelQText.Name = "labelQText";
            this.labelQText.Size = new System.Drawing.Size(28, 13);
            this.labelQText.TabIndex = 1;
            this.labelQText.Text = "Text";
            // 
            // labelQID
            // 
            this.labelQID.AutoSize = true;
            this.labelQID.Location = new System.Drawing.Point(17, 23);
            this.labelQID.Name = "labelQID";
            this.labelQID.Size = new System.Drawing.Size(21, 13);
            this.labelQID.TabIndex = 0;
            this.labelQID.Text = "ID:";
            // 
            // groupBoxA
            // 
            this.groupBoxA.Controls.Add(this.radioButtonMCNo);
            this.groupBoxA.Controls.Add(this.radioButtonMCYes);
            this.groupBoxA.Controls.Add(this.labelAType);
            this.groupBoxA.Location = new System.Drawing.Point(29, 187);
            this.groupBoxA.Name = "groupBoxA";
            this.groupBoxA.Size = new System.Drawing.Size(367, 93);
            this.groupBoxA.TabIndex = 1;
            this.groupBoxA.TabStop = false;
            this.groupBoxA.Text = "Answer type";
            // 
            // radioButtonMCNo
            // 
            this.radioButtonMCNo.AutoSize = true;
            this.radioButtonMCNo.Checked = true;
            this.radioButtonMCNo.Location = new System.Drawing.Point(113, 57);
            this.radioButtonMCNo.Name = "radioButtonMCNo";
            this.radioButtonMCNo.Size = new System.Drawing.Size(39, 17);
            this.radioButtonMCNo.TabIndex = 2;
            this.radioButtonMCNo.TabStop = true;
            this.radioButtonMCNo.Text = "No";
            this.radioButtonMCNo.UseVisualStyleBackColor = true;
            // 
            // radioButtonMCYes
            // 
            this.radioButtonMCYes.AutoSize = true;
            this.radioButtonMCYes.Location = new System.Drawing.Point(113, 33);
            this.radioButtonMCYes.Name = "radioButtonMCYes";
            this.radioButtonMCYes.Size = new System.Drawing.Size(43, 17);
            this.radioButtonMCYes.TabIndex = 1;
            this.radioButtonMCYes.TabStop = true;
            this.radioButtonMCYes.Text = "Yes";
            this.radioButtonMCYes.UseVisualStyleBackColor = true;
            // 
            // labelAType
            // 
            this.labelAType.AutoSize = true;
            this.labelAType.Location = new System.Drawing.Point(27, 33);
            this.labelAType.Name = "labelAType";
            this.labelAType.Size = new System.Drawing.Size(85, 13);
            this.labelAType.TabIndex = 0;
            this.labelAType.Text = "Multiple Choice?";
            // 
            // buttonNext
            // 
            this.buttonNext.Location = new System.Drawing.Point(104, 308);
            this.buttonNext.Name = "buttonNext";
            this.buttonNext.Size = new System.Drawing.Size(75, 23);
            this.buttonNext.TabIndex = 2;
            this.buttonNext.Text = "Next ..";
            this.buttonNext.UseVisualStyleBackColor = true;
            this.buttonNext.Click += new System.EventHandler(this.buttonNext_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.Location = new System.Drawing.Point(247, 308);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 3;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // groupBoxXPath
            // 
            this.groupBoxXPath.Controls.Add(this.textBoxXPath);
            this.groupBoxXPath.Location = new System.Drawing.Point(31, 15);
            this.groupBoxXPath.Name = "groupBoxXPath";
            this.groupBoxXPath.Size = new System.Drawing.Size(364, 53);
            this.groupBoxXPath.TabIndex = 4;
            this.groupBoxXPath.TabStop = false;
            this.groupBoxXPath.Text = "For XPath";
            // 
            // textBoxXPath
            // 
            this.textBoxXPath.Location = new System.Drawing.Point(17, 22);
            this.textBoxXPath.Name = "textBoxXPath";
            this.textBoxXPath.ReadOnly = true;
            this.textBoxXPath.Size = new System.Drawing.Size(333, 20);
            this.textBoxXPath.TabIndex = 0;
            // 
            // FormQuestion
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(425, 343);
            this.Controls.Add(this.groupBoxXPath);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonNext);
            this.Controls.Add(this.groupBoxA);
            this.Controls.Add(this.groupBoxQ);
            this.Name = "FormQuestion";
            this.Text = "Question Setup";
            this.groupBoxQ.ResumeLayout(false);
            this.groupBoxQ.PerformLayout();
            this.groupBoxA.ResumeLayout(false);
            this.groupBoxA.PerformLayout();
            this.groupBoxXPath.ResumeLayout(false);
            this.groupBoxXPath.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBoxQ;
        private System.Windows.Forms.GroupBox groupBoxA;
        private System.Windows.Forms.Button buttonNext;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.TextBox textBoxQText;
        public System.Windows.Forms.TextBox textBoxQID;
        private System.Windows.Forms.Label labelQText;
        private System.Windows.Forms.Label labelQID;
        private System.Windows.Forms.RadioButton radioButtonMCNo;
        private System.Windows.Forms.RadioButton radioButtonMCYes;
        private System.Windows.Forms.Label labelAType;
        private System.Windows.Forms.GroupBox groupBoxXPath;
        private System.Windows.Forms.TextBox textBoxXPath;
    }
}