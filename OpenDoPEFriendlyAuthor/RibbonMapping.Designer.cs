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
 *
 * Portions Copyright (c) Microsoft Corporation.  All rights reserved.
 * 
 * With respect to those portions (from http://xmlmapping.codeplex.com/license):
 * 

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
namespace XmlMappingTaskPane
{
    partial class RibbonMapping : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonMapping()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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


            this.tabOpenDoPEAuthorFriendly = this.Factory.CreateRibbonTab();
            this.groupMapping = this.Factory.CreateRibbonGroup();
            this.toggleButtonMapping = this.Factory.CreateRibbonToggleButton();
            this.buttonReplaceXML = this.Factory.CreateRibbonButton();
            this.groupControls = this.Factory.CreateRibbonGroup();
            this.buttonBind = this.Factory.CreateRibbonButton();
            this.buttonCondition = this.Factory.CreateRibbonButton();
            this.buttonRepeat = this.Factory.CreateRibbonButton();
            this.buttonEdit = this.Factory.CreateRibbonButton();
            this.buttonDelete = this.Factory.CreateRibbonButton();
            this.groupAdvanced = this.Factory.CreateRibbonGroup();
            this.toggleButtonDesignMode = this.Factory.CreateRibbonToggleButton();
            this.menuAdvanced = this.Factory.CreateRibbonMenu();
            this.buttonPartSelect = this.Factory.CreateRibbonButton();
            this.buttonXmlOptions = this.Factory.CreateRibbonButton();
            this.buttonClearAll = this.Factory.CreateRibbonButton();
            this.groupAbout = this.Factory.CreateRibbonGroup();
            this.buttonAbout = this.Factory.CreateRibbonButton();
            this.tabOpenDoPEAuthorFriendly.SuspendLayout();
            this.groupMapping.SuspendLayout();
            this.groupControls.SuspendLayout();
            this.groupAdvanced.SuspendLayout();
            this.groupAbout.SuspendLayout();
            // 
            // tabOpenDoPEAuthorFriendly
            // 
            this.tabOpenDoPEAuthorFriendly.Groups.Add(this.groupMapping);
            this.tabOpenDoPEAuthorFriendly.Groups.Add(this.groupControls);
            this.tabOpenDoPEAuthorFriendly.Groups.Add(this.groupAdvanced);
            this.tabOpenDoPEAuthorFriendly.Groups.Add(this.groupAbout);
            this.tabOpenDoPEAuthorFriendly.Label = System.Configuration.ConfigurationManager.AppSettings["MenuEntry"];
            this.tabOpenDoPEAuthorFriendly.Name = "tabOpenDoPEAuthorFriendly";
            // 
            // groupMapping
            // 
            this.groupMapping.Items.Add(this.toggleButtonMapping);
            this.groupMapping.Items.Add(this.buttonReplaceXML);
            this.groupMapping.Label = "Start";
            this.groupMapping.Name = "groupMapping";
            // 
            // toggleButtonMapping
            // 
            this.toggleButtonMapping.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.toggleButtonMapping.Image = global::XmlMappingTaskPane.Properties.Resources.RibbonIcon;
            this.toggleButtonMapping.KeyTip = "M";
            this.toggleButtonMapping.Label = "Show XML";
            this.toggleButtonMapping.Name = "toggleButtonMapping";
            this.toggleButtonMapping.ScreenTip = "Start/continue authoring";
            this.toggleButtonMapping.ShowImage = true;
            this.toggleButtonMapping.SuperTip = "Show/hide the task pane.  Will setup the document first, if necessary.";
            this.toggleButtonMapping.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonMapping_Click);
            // 
            // buttonReplaceXML
            // 
            this.buttonReplaceXML.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonReplaceXML.Label = "Replace XML";
            this.buttonReplaceXML.Name = "buttonReplaceXML";
            this.buttonReplaceXML.OfficeImageId = "ReviewTrackChanges";
            this.buttonReplaceXML.ScreenTip = "Swap existing XML sample";
            this.buttonReplaceXML.ShowImage = true;
            this.buttonReplaceXML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(buttonReplaceXML_Click);
            // 
            // groupControls
            // 
            this.groupControls.Items.Add(this.buttonBind);
            this.groupControls.Items.Add(this.buttonCondition);

            string repeatButtonEnabled = System.Configuration.ConfigurationManager.AppSettings["Ribbon.Button.Repeat.Enabled"];
            //if (repeatButtonEnabled.ToLower().Equals("true"))
            {
                this.groupControls.Items.Add(this.buttonRepeat);

                // 
                // buttonRepeat
                // 
                this.buttonRepeat.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
                this.buttonRepeat.Label = "Add Repeat";
                this.buttonRepeat.Name = "buttonRepeat";
                this.buttonRepeat.OfficeImageId = "OutlineShowDetail";
                this.buttonRepeat.ScreenTip = "Make selection repeat";
                this.buttonRepeat.ShowImage = true;
                this.buttonRepeat.SuperTip = "Wrap selection in repeat content control";
                this.buttonRepeat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonRepeat_Click);

            }
            this.groupControls.Items.Add(this.buttonEdit);
            this.groupControls.Items.Add(this.buttonDelete);
            this.groupControls.Label = "Control Structures";
            this.groupControls.Name = "groupControls";
            // 
            // buttonBind
            // 
            this.buttonBind.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonBind.Label = "Add Data Value";
            this.buttonBind.Name = "buttonBind";
            this.buttonBind.OfficeImageId = "AutoCorrect";
            this.buttonBind.ShowImage = true;
            this.buttonBind.SuperTip = "Use this if you want to manually enter an XPath.  Otherwise it is easier to drag/" +
    "drop.";
            this.buttonBind.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonBind_Click);
            // 
            // buttonCondition
            // 
            this.buttonCondition.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonCondition.Label = "Add Condition";
            this.buttonCondition.Name = "buttonCondition";
            this.buttonCondition.OfficeImageId = "MacroConditions";
            this.buttonCondition.ScreenTip = "Make selection conditional";
            this.buttonCondition.ShowImage = true;
            this.buttonCondition.SuperTip = "Wrap selection in conditional content control";
            this.buttonCondition.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCondition_Click);
            // 
            // buttonEdit
            // 
            this.buttonEdit.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonEdit.Label = "Edit";
            this.buttonEdit.Name = "buttonEdit";
            this.buttonEdit.OfficeImageId = "ReviewTrackChanges";
            this.buttonEdit.ScreenTip = "Edit this repeat or condition";
            this.buttonEdit.ShowImage = true;
            this.buttonEdit.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonEdit_Click);
            // 
            // buttonDelete
            // 
            this.buttonDelete.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonDelete.Label = "Remove control";
            this.buttonDelete.Name = "buttonDelete";
            this.buttonDelete.OfficeImageId = "WatermarkRemove";
            this.buttonDelete.ScreenTip = "Remove control (but keep contents)";
            this.buttonDelete.ShowImage = true;
            this.buttonDelete.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonDelete_Click);
            // 
            // groupAdvanced
            // 
            this.groupAdvanced.Items.Add(this.toggleButtonDesignMode);
            this.groupAdvanced.Items.Add(this.buttonClearAll);
            this.groupAdvanced.Label = "Advanced";
            this.groupAdvanced.Name = "groupAdvanced";
            // 
            // toggleButtonDesignMode
            // 
            this.toggleButtonDesignMode.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.toggleButtonDesignMode.Label = "Design Mode";
            this.toggleButtonDesignMode.Name = "toggleButtonDesignMode";
            this.toggleButtonDesignMode.OfficeImageId = "DesignMode";
            this.toggleButtonDesignMode.ShowImage = true;
            this.toggleButtonDesignMode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonDesignMode_Click);
            // 
            // menuAdvanced - only display it if one of its sub menus is relevant
            // 
            string xmlOptions = System.Configuration.ConfigurationManager.AppSettings["Ribbon.Button.XMLOptions.Enabled"];
            string switchPart = System.Configuration.ConfigurationManager.AppSettings["Ribbon.Button.SwitchPart.Enabled"];

            //if (switchPart.ToLower().Equals("true")
            //    || xmlOptions.ToLower().Equals("true") )
            {

                this.menuAdvanced.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
                this.menuAdvanced.Label = "Settings";
                this.menuAdvanced.Name = "menuAdvanced";
                this.menuAdvanced.OfficeImageId = "GroupTools";
                this.menuAdvanced.ShowImage = true;

                this.groupAdvanced.Items.Add(this.menuAdvanced);
            }
            // 
            // buttonPartSelect
            // 
            //if (switchPart.ToLower().Equals("true"))
            {
                this.buttonPartSelect.Label = "Switch XML part";
                this.buttonPartSelect.Name = "buttonPartSelect";
                this.buttonPartSelect.OfficeImageId = "ContentControlBuildingBlockGallery";
                this.buttonPartSelect.ShowImage = true;
                this.buttonPartSelect.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonPartSelect_Click);

                this.menuAdvanced.Items.Add(this.buttonPartSelect);
            }
            // 
            // buttonXmlOptions
            // 
            //if (xmlOptions.ToLower().Equals("true"))
            {
                this.buttonXmlOptions.Label = "XML Options";
                this.buttonXmlOptions.Name = "buttonXmlOptions";
                this.buttonXmlOptions.OfficeImageId = "GroupTools";
                this.buttonXmlOptions.ShowImage = true;
                this.buttonXmlOptions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonXmlOptions_Click);

                this.menuAdvanced.Items.Add(this.buttonXmlOptions);
            }

            // 
            // buttonClearAll
            // 
            this.buttonClearAll.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonClearAll.Label = "Remove mappings";
            this.buttonClearAll.Name = "buttonClearAll";
            this.buttonClearAll.OfficeImageId = "WatermarkRemove";
            this.buttonClearAll.ScreenTip = "Remove template functionality from this docx";
            this.buttonClearAll.ShowImage = true;
            this.buttonClearAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(buttonClearAll_Click);

            // 
            // groupAbout
            // 
            this.groupAbout.Items.Add(this.buttonAbout);
            this.groupAbout.Label = "About";
            this.groupAbout.Name = "groupAbout";
            // 
            // buttonAbout
            // 
            this.buttonAbout.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonAbout.Label = "About";
            this.buttonAbout.Name = "buttonAbout";
            this.buttonAbout.OfficeImageId = "Info";
            this.buttonAbout.ShowImage = true;
            this.buttonAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAbout_Click);
            // 
            // RibbonMapping
            // 
            this.Name = "RibbonMapping";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tabOpenDoPEAuthorFriendly);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonMapping_Load);
            this.tabOpenDoPEAuthorFriendly.ResumeLayout(false);
            this.tabOpenDoPEAuthorFriendly.PerformLayout();
            this.groupMapping.ResumeLayout(false);
            this.groupMapping.PerformLayout();
            this.groupControls.ResumeLayout(false);
            this.groupControls.PerformLayout();
            this.groupAdvanced.ResumeLayout(false);
            this.groupAdvanced.PerformLayout();
            this.groupAbout.ResumeLayout(false);
            this.groupAbout.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabOpenDoPEAuthorFriendly;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupMapping;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonMapping;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupControls;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCondition;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonRepeat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonEdit;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAdvanced;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuAdvanced;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonPartSelect;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonDesignMode;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonXmlOptions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonBind;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonDelete;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonClearAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonReplaceXML;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonMapping RibbonMapping
        {
            get { return this.GetRibbon<RibbonMapping>(); }
        }
    }
}
