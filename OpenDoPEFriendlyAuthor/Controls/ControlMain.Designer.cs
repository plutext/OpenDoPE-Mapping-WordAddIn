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
namespace XmlMappingTaskPane.Controls
{
    partial class ControlMain
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
            this.controlTreeView = new XmlMappingTaskPane.Controls.ControlTreeView();
            this.splitContainer = new System.Windows.Forms.SplitContainer();
            this.controlProperties = new XmlMappingTaskPane.Controls.ControlProperties();
            this.controlMode1 = new XmlMappingTaskPane.Controls.ControlMode();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer)).BeginInit();
            this.splitContainer.Panel1.SuspendLayout();
            this.splitContainer.Panel2.SuspendLayout();
            this.splitContainer.SuspendLayout();
            this.SuspendLayout();
            // 
            // controlTreeView
            // 
            this.controlTreeView.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.controlTreeView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.controlTreeView.Location = new System.Drawing.Point(0, 0);
            this.controlTreeView.Name = "controlTreeView";
            this.controlTreeView.Size = new System.Drawing.Size(242, 220);
            this.controlTreeView.TabIndex = 1;
            this.controlTreeView.Load += new System.EventHandler(this.controlTreeView_Load);
            // 
            // splitContainer (contains the tree view & properties)
            // 
            this.splitContainer.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            if (modeControlEnabled)
            {
                this.splitContainer.Location = new System.Drawing.Point(3, 80);
            }
            else
            {
                this.splitContainer.Location = new System.Drawing.Point(3, 3);
            }
            this.splitContainer.Name = "splitContainer";
            this.splitContainer.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer.Panel1
            // 
            this.splitContainer.Panel1.Controls.Add(this.controlTreeView);
            // 
            // splitContainer.Panel2
            // 
            this.splitContainer.Panel2.Controls.Add(this.controlProperties);
            this.splitContainer.Size = new System.Drawing.Size(242, 303);
            this.splitContainer.SplitterDistance = 220;
            this.splitContainer.TabIndex = 2;
            // 
            // controlProperties
            // 
            this.controlProperties.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.controlProperties.Location = new System.Drawing.Point(0, 2);
            this.controlProperties.Name = "controlProperties";
            this.controlProperties.Size = new System.Drawing.Size(242, 77);
            this.controlProperties.TabIndex = 0;
            // 
            // controlMode1
            // 
            if (modeControlEnabled)
            {
                this.controlMode1.Location = new System.Drawing.Point(3, 3);
                this.controlMode1.Name = "controlMode1";
                this.controlMode1.Size = new System.Drawing.Size(242, 74);
                this.controlMode1.TabIndex = 3;
            }
            // 
            // ControlMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            if (modeControlEnabled)
            {
                this.Controls.Add(this.controlMode1);
            }
            this.Controls.Add(this.splitContainer);
            this.Name = "ControlMain";
            this.Size = new System.Drawing.Size(248, 433);
            this.Resize += new System.EventHandler(this.ControlMain_Resize);
            this.splitContainer.Panel1.ResumeLayout(false);
            this.splitContainer.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer)).EndInit();
            this.splitContainer.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        public ControlTreeView controlTreeView { get; set; }
        private System.Windows.Forms.SplitContainer splitContainer;
        private ControlProperties controlProperties;
        public ControlMode controlMode1 { get; set; }
    }
}
