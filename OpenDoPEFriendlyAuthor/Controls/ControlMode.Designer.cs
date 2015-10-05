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
namespace XmlMappingTaskPane.Controls
{
    partial class ControlMode
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
            this.radioModeBind = new System.Windows.Forms.RadioButton();
            this.radioModeCondition = new System.Windows.Forms.RadioButton();
            this.radioModeRepeat = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // radioModeBind
            // 
            this.radioModeBind.Appearance = System.Windows.Forms.Appearance.Button;
            this.radioModeBind.AutoSize = true;
            this.radioModeBind.Checked = true;
            this.radioModeBind.Location = new System.Drawing.Point(6, 21);
            this.radioModeBind.Name = "radioModeBind";
            this.radioModeBind.Size = new System.Drawing.Size(69, 23);
            this.radioModeBind.TabIndex = 1;
            this.radioModeBind.TabStop = true;
            this.radioModeBind.Text = "Data value";
            this.radioModeBind.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.radioModeBind.UseVisualStyleBackColor = true;
            // 
            // radioModeCondition
            // 
            this.radioModeCondition.Appearance = System.Windows.Forms.Appearance.Button;
            this.radioModeCondition.AutoSize = true;
            this.radioModeCondition.Location = new System.Drawing.Point(81, 22);
            this.radioModeCondition.MinimumSize = new System.Drawing.Size(69, 0);
            this.radioModeCondition.Name = "radioModeCondition";
            this.radioModeCondition.Size = new System.Drawing.Size(69, 23);
            this.radioModeCondition.TabIndex = 2;
            this.radioModeCondition.TabStop = true;
            this.radioModeCondition.Text = "Condition";
            this.radioModeCondition.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.radioModeCondition.UseVisualStyleBackColor = true;
            // 
            // radioModeRepeat
            // 
            this.radioModeRepeat.Appearance = System.Windows.Forms.Appearance.Button;
            this.radioModeRepeat.AutoSize = true;
            this.radioModeRepeat.Location = new System.Drawing.Point(156, 22);
            this.radioModeRepeat.MinimumSize = new System.Drawing.Size(69, 0);
            this.radioModeRepeat.Name = "radioModeRepeat";
            this.radioModeRepeat.Size = new System.Drawing.Size(69, 23);
            this.radioModeRepeat.TabIndex = 3;
            this.radioModeRepeat.Text = "Repeat";
            this.radioModeRepeat.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.radioModeRepeat.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radioModeBind);
            this.groupBox1.Controls.Add(this.radioModeRepeat);
            this.groupBox1.Controls.Add(this.radioModeCondition);
            this.groupBox1.Location = new System.Drawing.Point(7, 11);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(232, 54);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Type";
            // 
            // ControlMode
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.groupBox1);
            this.Name = "ControlMode";
            this.Size = new System.Drawing.Size(242, 82);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.RadioButton radioModeBind;
        private System.Windows.Forms.RadioButton radioModeCondition;
        private System.Windows.Forms.RadioButton radioModeRepeat;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}
