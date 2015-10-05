namespace XmlMappingTaskPane.Controls
{
    partial class ControlPartList
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ControlPartList));
            this.contextMenuPart = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.renameToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.deleteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolTipNamespace = new System.Windows.Forms.ToolTip(this.components);
            this.labelDataSource = new System.Windows.Forms.Label();
            this.comboBoxPartList = new System.Windows.Forms.ComboBox();
            this.contextMenuPart.SuspendLayout();
            this.SuspendLayout();
            // 
            // contextMenuPart
            // 
            resources.ApplyResources(this.contextMenuPart, "contextMenuPart");
            this.contextMenuPart.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.renameToolStripMenuItem,
            this.deleteToolStripMenuItem});
            this.contextMenuPart.Name = "contextMenuPart";
            this.contextMenuPart.Opening += new System.ComponentModel.CancelEventHandler(this.contextMenuPart_Opening);
            // 
            // renameToolStripMenuItem
            // 
            this.renameToolStripMenuItem.Name = "renameToolStripMenuItem";
            resources.ApplyResources(this.renameToolStripMenuItem, "renameToolStripMenuItem");
            this.renameToolStripMenuItem.Click += new System.EventHandler(this.renameToolStripMenuItem_Click);
            // 
            // deleteToolStripMenuItem
            // 
            this.deleteToolStripMenuItem.Name = "deleteToolStripMenuItem";
            resources.ApplyResources(this.deleteToolStripMenuItem, "deleteToolStripMenuItem");
            this.deleteToolStripMenuItem.Click += new System.EventHandler(this.deleteToolStripMenuItem_Click);
            // 
            // labelDataSource
            // 
            resources.ApplyResources(this.labelDataSource, "labelDataSource");
            this.labelDataSource.Name = "labelDataSource";
            // 
            // comboBoxPartList
            // 
            resources.ApplyResources(this.comboBoxPartList, "comboBoxPartList");
            this.comboBoxPartList.ContextMenuStrip = this.contextMenuPart;
            this.comboBoxPartList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxPartList.FormattingEnabled = true;
            this.comboBoxPartList.Name = "comboBoxPartList";
            this.comboBoxPartList.SelectedIndexChanged += new System.EventHandler(this.comboBoxPartList_SelectedIndexChanged);
            this.comboBoxPartList.MouseHover += new System.EventHandler(this.comboBoxPartList_MouseHover);
            // 
            // ControlPartList
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.comboBoxPartList);
            this.Controls.Add(this.labelDataSource);
            this.Name = "ControlPartList";
            this.contextMenuPart.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label labelDataSource;
        private System.Windows.Forms.ComboBox comboBoxPartList;
        private System.Windows.Forms.ContextMenuStrip contextMenuPart;
        private System.Windows.Forms.ToolTip toolTipNamespace;
        private System.Windows.Forms.ToolStripMenuItem renameToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem deleteToolStripMenuItem;
    }
}
