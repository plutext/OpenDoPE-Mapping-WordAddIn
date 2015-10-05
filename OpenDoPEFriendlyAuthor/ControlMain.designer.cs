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
            this.controlPartList = new XmlMappingTaskPane.Controls.ControlPartList();
            this.splitContainer = new System.Windows.Forms.SplitContainer();
            this.controlProperties = new XmlMappingTaskPane.Controls.ControlProperties();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer)).BeginInit();
            this.splitContainer.Panel1.SuspendLayout();
            this.splitContainer.Panel2.SuspendLayout();
            this.splitContainer.SuspendLayout();
            this.SuspendLayout();
            // 
            // controlTreeView
            // 
            this.controlTreeView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.controlTreeView.Location = new System.Drawing.Point(0, 0);
            this.controlTreeView.Name = "controlTreeView";
            this.controlTreeView.Size = new System.Drawing.Size(242, 279);
            this.controlTreeView.TabIndex = 1;
            // 
            // controlPartList
            // 
            this.controlPartList.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.controlPartList.Location = new System.Drawing.Point(0, 0);
            this.controlPartList.Name = "controlPartList";
            this.controlPartList.Size = new System.Drawing.Size(248, 47);
            this.controlPartList.TabIndex = 0;
            // 
            // splitContainer
            // 
            this.splitContainer.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.splitContainer.Location = new System.Drawing.Point(3, 46);
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
            this.splitContainer.Size = new System.Drawing.Size(242, 384);
            this.splitContainer.SplitterDistance = 279;
            this.splitContainer.TabIndex = 2;
            // 
            // controlProperties
            // 
            this.controlProperties.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.controlProperties.Location = new System.Drawing.Point(0, 2);
            this.controlProperties.Name = "controlProperties";
            this.controlProperties.Size = new System.Drawing.Size(242, 99);
            this.controlProperties.TabIndex = 0;
            // 
            // ControlMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.splitContainer);
            this.Controls.Add(this.controlPartList);
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

        private ControlPartList controlPartList;
        private ControlTreeView controlTreeView;
        private System.Windows.Forms.SplitContainer splitContainer;
        private ControlProperties controlProperties;
    }
}
