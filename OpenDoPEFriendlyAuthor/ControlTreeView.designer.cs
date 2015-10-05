namespace XmlMappingTaskPane.Controls
{
    partial class ControlTreeView
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ControlTreeView));
            this.backgroundWorkerMain = new System.ComponentModel.BackgroundWorker();
            this.backgroundWorkerBuildTree = new System.ComponentModel.BackgroundWorker();
            this.timerLoading = new System.Windows.Forms.Timer(this.components);
            this.labelLoading = new System.Windows.Forms.Label();
            this.treeView = new System.Windows.Forms.TreeView();
            this.contextMenuNode = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.mapToSelectedControlToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.insertToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.textDefaultToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dateDefaultToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dropDownListDefaultToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pictureDefaultToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator = new System.Windows.Forms.ToolStripSeparator();
            this.comboBoxToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dateToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dropDownListToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pictureToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.textToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.addToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.belowToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.insideToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.deleteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.imageList = new System.Windows.Forms.ImageList(this.components);
            this.contextMenuNode.SuspendLayout();
            this.SuspendLayout();
            // 
            // backgroundWorkerMain
            // 
            this.backgroundWorkerMain.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorkerMain_DoWork);
            // 
            // backgroundWorkerBuildTree
            // 
            this.backgroundWorkerBuildTree.WorkerReportsProgress = true;
            this.backgroundWorkerBuildTree.WorkerSupportsCancellation = true;
            this.backgroundWorkerBuildTree.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorkerBuildTree_DoWork);
            this.backgroundWorkerBuildTree.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorkerBuildTree_ProgressChanged);
            // 
            // timerLoading
            // 
            this.timerLoading.Interval = 2000;
            this.timerLoading.Tick += new System.EventHandler(this.timerLoading_Tick);
            // 
            // labelLoading
            // 
            resources.ApplyResources(this.labelLoading, "labelLoading");
            this.labelLoading.BackColor = System.Drawing.SystemColors.Window;
            this.labelLoading.Name = "labelLoading";
            this.labelLoading.UseWaitCursor = true;
            // 
            // treeView
            // 
            this.treeView.AllowDrop = true;
            this.treeView.ContextMenuStrip = this.contextMenuNode;
            resources.ApplyResources(this.treeView, "treeView");
            this.treeView.HideSelection = false;
            this.treeView.ImageList = this.imageList;
            this.treeView.Name = "treeView";
            this.treeView.ShowLines = false;
            this.treeView.AfterLabelEdit += new System.Windows.Forms.NodeLabelEditEventHandler(this.treeView_AfterLabelEdit);
            this.treeView.ItemDrag += new System.Windows.Forms.ItemDragEventHandler(this.treeView_ItemDrag);
            this.treeView.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeView_AfterSelect);
            this.treeView.NodeMouseDoubleClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.treeView_NodeMouseDoubleClick);
            this.treeView.DragDrop += new System.Windows.Forms.DragEventHandler(this.treeView_DragDrop);
            this.treeView.DragEnter += new System.Windows.Forms.DragEventHandler(this.treeView_DragEnter);
            this.treeView.MouseDown += new System.Windows.Forms.MouseEventHandler(this.treeView_MouseDown);
            // 
            // contextMenuNode
            // 
            resources.ApplyResources(this.contextMenuNode, "contextMenuNode");
            this.contextMenuNode.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mapToSelectedControlToolStripMenuItem,
            this.insertToolStripMenuItem,
            this.toolStripSeparator1,
            this.addToolStripMenuItem,
            this.deleteToolStripMenuItem});
            this.contextMenuNode.Name = "contextMenuStrip1";
            this.contextMenuNode.Opening += new System.ComponentModel.CancelEventHandler(this.contextMenuNode_Opening);
            // 
            // mapToSelectedControlToolStripMenuItem
            // 
            this.mapToSelectedControlToolStripMenuItem.Name = "mapToSelectedControlToolStripMenuItem";
            resources.ApplyResources(this.mapToSelectedControlToolStripMenuItem, "mapToSelectedControlToolStripMenuItem");
            this.mapToSelectedControlToolStripMenuItem.Click += new System.EventHandler(this.mapToSelectedControlToolStripMenuItem_Click);
            // 
            // insertToolStripMenuItem
            // 
            this.insertToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.textDefaultToolStripMenuItem,
            this.dateDefaultToolStripMenuItem,
            this.dropDownListDefaultToolStripMenuItem,
            this.pictureDefaultToolStripMenuItem,
            this.toolStripSeparator,
            this.comboBoxToolStripMenuItem,
            this.dateToolStripMenuItem,
            this.dropDownListToolStripMenuItem,
            this.pictureToolStripMenuItem,
            this.textToolStripMenuItem});
            this.insertToolStripMenuItem.Name = "insertToolStripMenuItem";
            resources.ApplyResources(this.insertToolStripMenuItem, "insertToolStripMenuItem");
            // 
            // textDefaultToolStripMenuItem
            // 
            this.textDefaultToolStripMenuItem.Name = "textDefaultToolStripMenuItem";
            resources.ApplyResources(this.textDefaultToolStripMenuItem, "textDefaultToolStripMenuItem");
            this.textDefaultToolStripMenuItem.Click += new System.EventHandler(this.textToolStripMenuItem_Click);
            // 
            // dateDefaultToolStripMenuItem
            // 
            this.dateDefaultToolStripMenuItem.Name = "dateDefaultToolStripMenuItem";
            resources.ApplyResources(this.dateDefaultToolStripMenuItem, "dateDefaultToolStripMenuItem");
            this.dateDefaultToolStripMenuItem.Click += new System.EventHandler(this.dateToolStripMenuItem_Click);
            // 
            // dropDownListDefaultToolStripMenuItem
            // 
            this.dropDownListDefaultToolStripMenuItem.Name = "dropDownListDefaultToolStripMenuItem";
            resources.ApplyResources(this.dropDownListDefaultToolStripMenuItem, "dropDownListDefaultToolStripMenuItem");
            this.dropDownListDefaultToolStripMenuItem.Click += new System.EventHandler(this.dropDownListToolStripMenuItem_Click);
            // 
            // pictureDefaultToolStripMenuItem
            // 
            this.pictureDefaultToolStripMenuItem.Name = "pictureDefaultToolStripMenuItem";
            resources.ApplyResources(this.pictureDefaultToolStripMenuItem, "pictureDefaultToolStripMenuItem");
            this.pictureDefaultToolStripMenuItem.Click += new System.EventHandler(this.pictureToolStripMenuItem_Click);
            // 
            // toolStripSeparator
            // 
            this.toolStripSeparator.Name = "toolStripSeparator";
            resources.ApplyResources(this.toolStripSeparator, "toolStripSeparator");
            // 
            // comboBoxToolStripMenuItem
            // 
            this.comboBoxToolStripMenuItem.Name = "comboBoxToolStripMenuItem";
            resources.ApplyResources(this.comboBoxToolStripMenuItem, "comboBoxToolStripMenuItem");
            this.comboBoxToolStripMenuItem.Click += new System.EventHandler(this.comboBoxToolStripMenuItem_Click);
            // 
            // dateToolStripMenuItem
            // 
            this.dateToolStripMenuItem.Name = "dateToolStripMenuItem";
            resources.ApplyResources(this.dateToolStripMenuItem, "dateToolStripMenuItem");
            this.dateToolStripMenuItem.Click += new System.EventHandler(this.dateToolStripMenuItem_Click);
            // 
            // dropDownListToolStripMenuItem
            // 
            this.dropDownListToolStripMenuItem.Name = "dropDownListToolStripMenuItem";
            resources.ApplyResources(this.dropDownListToolStripMenuItem, "dropDownListToolStripMenuItem");
            this.dropDownListToolStripMenuItem.Click += new System.EventHandler(this.dropDownListToolStripMenuItem_Click);
            // 
            // pictureToolStripMenuItem
            // 
            this.pictureToolStripMenuItem.Name = "pictureToolStripMenuItem";
            resources.ApplyResources(this.pictureToolStripMenuItem, "pictureToolStripMenuItem");
            this.pictureToolStripMenuItem.Click += new System.EventHandler(this.pictureToolStripMenuItem_Click);
            // 
            // textToolStripMenuItem
            // 
            this.textToolStripMenuItem.Name = "textToolStripMenuItem";
            resources.ApplyResources(this.textToolStripMenuItem, "textToolStripMenuItem");
            this.textToolStripMenuItem.Click += new System.EventHandler(this.textToolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            resources.ApplyResources(this.toolStripSeparator1, "toolStripSeparator1");
            // 
            // addToolStripMenuItem
            // 
            this.addToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.aboveToolStripMenuItem,
            this.belowToolStripMenuItem,
            this.insideToolStripMenuItem});
            this.addToolStripMenuItem.Name = "addToolStripMenuItem";
            resources.ApplyResources(this.addToolStripMenuItem, "addToolStripMenuItem");
            // 
            // aboveToolStripMenuItem
            // 
            this.aboveToolStripMenuItem.Name = "aboveToolStripMenuItem";
            resources.ApplyResources(this.aboveToolStripMenuItem, "aboveToolStripMenuItem");
            this.aboveToolStripMenuItem.Click += new System.EventHandler(this.aboveToolStripMenuItem_Click);
            // 
            // belowToolStripMenuItem
            // 
            this.belowToolStripMenuItem.Name = "belowToolStripMenuItem";
            resources.ApplyResources(this.belowToolStripMenuItem, "belowToolStripMenuItem");
            this.belowToolStripMenuItem.Click += new System.EventHandler(this.belowToolStripMenuItem_Click);
            // 
            // insideToolStripMenuItem
            // 
            this.insideToolStripMenuItem.Name = "insideToolStripMenuItem";
            resources.ApplyResources(this.insideToolStripMenuItem, "insideToolStripMenuItem");
            this.insideToolStripMenuItem.Click += new System.EventHandler(this.insideToolStripMenuItem_Click);
            // 
            // deleteToolStripMenuItem
            // 
            this.deleteToolStripMenuItem.Name = "deleteToolStripMenuItem";
            resources.ApplyResources(this.deleteToolStripMenuItem, "deleteToolStripMenuItem");
            this.deleteToolStripMenuItem.Click += new System.EventHandler(this.deleteToolStripMenuItem_Click);
            // 
            // imageList
            // 
            this.imageList.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList.ImageStream")));
            this.imageList.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList.Images.SetKeyName(0, "Group.png");
            this.imageList.Images.SetKeyName(1, "Text.png");
            this.imageList.Images.SetKeyName(2, "Element.png");
            this.imageList.Images.SetKeyName(3, "Attribute.png");
            // 
            // ControlTreeView
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.labelLoading);
            this.Controls.Add(this.treeView);
            this.Name = "ControlTreeView";
            this.contextMenuNode.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.ComponentModel.BackgroundWorker backgroundWorkerMain;
        private System.ComponentModel.BackgroundWorker backgroundWorkerBuildTree;
        private System.Windows.Forms.Timer timerLoading;
        private System.Windows.Forms.TreeView treeView;
        private System.Windows.Forms.Label labelLoading;
        private System.Windows.Forms.ContextMenuStrip contextMenuNode;
        private System.Windows.Forms.ToolStripMenuItem mapToSelectedControlToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem insertToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem textDefaultToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem dateDefaultToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem dropDownListDefaultToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem pictureDefaultToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator;
        private System.Windows.Forms.ToolStripMenuItem textToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem dateToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem dropDownListToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem pictureToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem addToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboveToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem belowToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem insideToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem deleteToolStripMenuItem;
        private System.Windows.Forms.ImageList imageList;
        private System.Windows.Forms.ToolStripMenuItem comboBoxToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
    }
}
