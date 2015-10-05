/*
 *  OpenDoPE authoring Word AddIn
    Copyright (C) Plutext Pty Ltd, 2012
 * 
    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */
namespace XmlMappingTaskPane.Forms
{
    partial class FormSwitchSelectedPart
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
            this.controlPartList = new XmlMappingTaskPane.Controls.ControlPartList();
            this.buttonHide = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // controlPartList1
            // 
            this.controlPartList.Location = new System.Drawing.Point(12, 21);
            this.controlPartList.Name = "controlPartList";
            this.controlPartList.Size = new System.Drawing.Size(243, 43);
            this.controlPartList.TabIndex = 0;
            // 
            // buttonHide
            // 
            this.buttonHide.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonHide.Location = new System.Drawing.Point(180, 88);
            this.buttonHide.Name = "buttonHide";
            this.buttonHide.Size = new System.Drawing.Size(75, 23);
            this.buttonHide.TabIndex = 2;
            this.buttonHide.Text = "Hide";
            this.buttonHide.UseVisualStyleBackColor = true;
            this.buttonHide.Click +=new System.EventHandler(buttonHide_Click);
            // 
            // FormSwitchSelectedPart
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(268, 125);
            this.Controls.Add(this.buttonHide);
            this.Controls.Add(this.controlPartList);
            this.Name = "FormSwitchSelectedPart";
            this.Text = "Select XML part";
            this.ResumeLayout(false);
            this.FormClosing +=new System.Windows.Forms.FormClosingEventHandler(FormSwitchSelectedPart_FormClosing);

        }

        #endregion

        public Controls.ControlPartList controlPartList { get; set; }
        private System.Windows.Forms.Button buttonHide;

    }
}