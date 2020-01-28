namespace MDL
{
    partial class ProcessTool
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
            //base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ProcessTool));
            this.defaultLookAndFeel1 = new DevExpress.LookAndFeel.DefaultLookAndFeel(this.components);
            this.ribbonControl1 = new DevExpress.XtraBars.Ribbon.RibbonControl();
            this.spreadsheetCommandBarButtonItem12 = new DevExpress.XtraSpreadsheet.UI.SpreadsheetCommandBarButtonItem();
            this.saveButton = new DevExpress.XtraBars.BarButtonItem();
            this.btnHelp = new DevExpress.XtraBars.BarButtonItem();
            this.btnExport = new DevExpress.XtraBars.BarButtonItem();
            this.fileRibbonPage1 = new DevExpress.XtraSpreadsheet.UI.FileRibbonPage();
            this.commonRibbonPageGroup4 = new DevExpress.XtraSpreadsheet.UI.CommonRibbonPageGroup();
            this.Inforbn = new DevExpress.XtraBars.Ribbon.RibbonPageGroup();
            this.pnlExcelctrl = new DevExpress.XtraEditors.PanelControl();
            this.spreadsheetControl1 = new DevExpress.XtraSpreadsheet.SpreadsheetControl();
            this.commonRibbonPageGroup1 = new DevExpress.XtraSpreadsheet.UI.CommonRibbonPageGroup();
            this.commonRibbonPageGroup2 = new DevExpress.XtraSpreadsheet.UI.CommonRibbonPageGroup();
            this.commonRibbonPageGroup3 = new DevExpress.XtraSpreadsheet.UI.CommonRibbonPageGroup();
            this.spreadsheetCommandBarButtonItem1 = new DevExpress.XtraSpreadsheet.UI.SpreadsheetCommandBarButtonItem();
            this.spreadsheetBarController1 = new DevExpress.XtraSpreadsheet.UI.SpreadsheetBarController(this.components);
            this.spreadsheetCommandBarButtonItem6 = new DevExpress.XtraSpreadsheet.UI.SpreadsheetCommandBarButtonItem();
            this.spreadsheetCommandBarButtonItem5 = new DevExpress.XtraSpreadsheet.UI.SpreadsheetCommandBarButtonItem();
            this.spreadsheetCommandBarButtonItem2 = new DevExpress.XtraSpreadsheet.UI.SpreadsheetCommandBarButtonItem();
            ((System.ComponentModel.ISupportInitialize)(this.ribbonControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pnlExcelctrl)).BeginInit();
            this.pnlExcelctrl.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.spreadsheetBarController1)).BeginInit();
            this.SuspendLayout();
            // 
            // defaultLookAndFeel1
            // 
            this.defaultLookAndFeel1.LookAndFeel.SkinName = "MySkin_Money Twins1";
            // 
            // ribbonControl1
            // 
            this.ribbonControl1.ExpandCollapseItem.Id = 0;
            this.ribbonControl1.Items.AddRange(new DevExpress.XtraBars.BarItem[] {
            this.ribbonControl1.ExpandCollapseItem,
            this.ribbonControl1.SearchEditItem,
            this.spreadsheetCommandBarButtonItem12,
            this.saveButton,
            this.btnHelp,
            this.btnExport});
            this.ribbonControl1.Location = new System.Drawing.Point(0, 0);
            this.ribbonControl1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.ribbonControl1.MaxItemId = 21;
            this.ribbonControl1.Name = "ribbonControl1";
            this.ribbonControl1.Pages.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPage[] {
            this.fileRibbonPage1});
            this.ribbonControl1.Size = new System.Drawing.Size(1195, 183);
            // 
            // spreadsheetCommandBarButtonItem12
            // 
            this.spreadsheetCommandBarButtonItem12.CommandName = "FileShowDocumentProperties";
            this.spreadsheetCommandBarButtonItem12.Id = 11;
            this.spreadsheetCommandBarButtonItem12.Name = "spreadsheetCommandBarButtonItem12";
            // 
            // saveButton
            // 
            this.saveButton.Caption = "Save";
            this.saveButton.Id = 12;
            this.saveButton.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("saveButton.ImageOptions.Image")));
            this.saveButton.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("saveButton.ImageOptions.LargeImage")));
            this.saveButton.Name = "saveButton";
            this.saveButton.RibbonStyle = ((DevExpress.XtraBars.Ribbon.RibbonItemStyles)(((DevExpress.XtraBars.Ribbon.RibbonItemStyles.Large | DevExpress.XtraBars.Ribbon.RibbonItemStyles.SmallWithText) 
            | DevExpress.XtraBars.Ribbon.RibbonItemStyles.SmallWithoutText)));
            this.saveButton.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.saveButton_ItemClick);
            // 
            // btnHelp
            // 
            this.btnHelp.Caption = "Help";
            this.btnHelp.Id = 16;
            this.btnHelp.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("btnHelp.ImageOptions.LargeImage")));
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.RibbonStyle = ((DevExpress.XtraBars.Ribbon.RibbonItemStyles)(((DevExpress.XtraBars.Ribbon.RibbonItemStyles.Large | DevExpress.XtraBars.Ribbon.RibbonItemStyles.SmallWithText) 
            | DevExpress.XtraBars.Ribbon.RibbonItemStyles.SmallWithoutText)));
            this.btnHelp.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btnHelp_ItemClick);
            // 
            // btnExport
            // 
            this.btnExport.Caption = "Export";
            this.btnExport.Id = 20;
            this.btnExport.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btnExport.ImageOptions.Image")));
            this.btnExport.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("btnExport.ImageOptions.LargeImage")));
            this.btnExport.Name = "btnExport";
            this.btnExport.RibbonStyle = ((DevExpress.XtraBars.Ribbon.RibbonItemStyles)(((DevExpress.XtraBars.Ribbon.RibbonItemStyles.Large | DevExpress.XtraBars.Ribbon.RibbonItemStyles.SmallWithText) 
            | DevExpress.XtraBars.Ribbon.RibbonItemStyles.SmallWithoutText)));
            this.btnExport.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btnExport_ItemClick);
            // 
            // fileRibbonPage1
            // 
            this.fileRibbonPage1.Groups.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPageGroup[] {
            this.commonRibbonPageGroup4,
            this.Inforbn});
            this.fileRibbonPage1.Name = "fileRibbonPage1";
            // 
            // commonRibbonPageGroup4
            // 
            this.commonRibbonPageGroup4.CaptionButtonVisible = DevExpress.Utils.DefaultBoolean.False;
            this.commonRibbonPageGroup4.ItemLinks.Add(this.saveButton);
            this.commonRibbonPageGroup4.ItemLinks.Add(this.btnExport);
            this.commonRibbonPageGroup4.Name = "commonRibbonPageGroup4";
            this.commonRibbonPageGroup4.Text = "";
            // 
            // Inforbn
            // 
            this.Inforbn.ItemLinks.Add(this.btnHelp);
            this.Inforbn.Name = "Inforbn";
            this.Inforbn.Text = "Info";
            // 
            // pnlExcelctrl
            // 
            this.pnlExcelctrl.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlExcelctrl.Controls.Add(this.spreadsheetControl1);
            this.pnlExcelctrl.Location = new System.Drawing.Point(7, 190);
            this.pnlExcelctrl.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.pnlExcelctrl.Name = "pnlExcelctrl";
            this.pnlExcelctrl.Size = new System.Drawing.Size(1175, 500);
            this.pnlExcelctrl.TabIndex = 11;
            // 
            // spreadsheetControl1
            // 
            this.spreadsheetControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.spreadsheetControl1.Location = new System.Drawing.Point(5, 5);
            this.spreadsheetControl1.LookAndFeel.UseDefaultLookAndFeel = false;
            this.spreadsheetControl1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.spreadsheetControl1.MenuManager = this.ribbonControl1;
            this.spreadsheetControl1.Name = "spreadsheetControl1";
            this.spreadsheetControl1.Options.Behavior.Worksheet.Delete = DevExpress.XtraSpreadsheet.DocumentCapability.Disabled;
            this.spreadsheetControl1.Options.Behavior.Worksheet.Rename = DevExpress.XtraSpreadsheet.DocumentCapability.Disabled;
            this.spreadsheetControl1.Size = new System.Drawing.Size(1165, 490);
            this.spreadsheetControl1.TabIndex = 0;
            this.spreadsheetControl1.Text = "spreadsheetControl1";
            this.spreadsheetControl1.CellValueChanged += new DevExpress.XtraSpreadsheet.CellValueChangedEventHandler(this.spreadsheetControl1_CellValueChanged);
            this.spreadsheetControl1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.spreadsheetControl1_KeyDown);
            // 
            // commonRibbonPageGroup1
            // 
            this.commonRibbonPageGroup1.CaptionButtonVisible = DevExpress.Utils.DefaultBoolean.False;
            this.commonRibbonPageGroup1.Name = "commonRibbonPageGroup1";
            this.commonRibbonPageGroup1.Text = "";
            // 
            // commonRibbonPageGroup2
            // 
            this.commonRibbonPageGroup2.CaptionButtonVisible = DevExpress.Utils.DefaultBoolean.False;
            this.commonRibbonPageGroup2.Name = "commonRibbonPageGroup2";
            this.commonRibbonPageGroup2.Text = "";
            // 
            // commonRibbonPageGroup3
            // 
            this.commonRibbonPageGroup3.CaptionButtonVisible = DevExpress.Utils.DefaultBoolean.False;
            this.commonRibbonPageGroup3.Name = "commonRibbonPageGroup3";
            this.commonRibbonPageGroup3.Text = "";
            // 
            // spreadsheetCommandBarButtonItem1
            // 
            this.spreadsheetCommandBarButtonItem1.CommandName = "FileNew";
            this.spreadsheetCommandBarButtonItem1.Id = 1;
            this.spreadsheetCommandBarButtonItem1.Name = "spreadsheetCommandBarButtonItem1";
            // 
            // spreadsheetBarController1
            // 
            this.spreadsheetBarController1.BarItems.Add(this.spreadsheetCommandBarButtonItem12);
            this.spreadsheetBarController1.Control = this.spreadsheetControl1;
            // 
            // spreadsheetCommandBarButtonItem6
            // 
            this.spreadsheetCommandBarButtonItem6.Name = "spreadsheetCommandBarButtonItem6";
            // 
            // spreadsheetCommandBarButtonItem5
            // 
            this.spreadsheetCommandBarButtonItem5.CommandName = "FileSaveAs";
            this.spreadsheetCommandBarButtonItem5.Id = 4;
            this.spreadsheetCommandBarButtonItem5.Name = "spreadsheetCommandBarButtonItem5";
            // 
            // spreadsheetCommandBarButtonItem2
            // 
            this.spreadsheetCommandBarButtonItem2.Name = "spreadsheetCommandBarButtonItem2";
            // 
            // ProcessTool
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1195, 702);
            this.Controls.Add(this.pnlExcelctrl);
            this.Controls.Add(this.ribbonControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.Name = "ProcessTool";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "MDL Input Setting File Process";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ProcessTool_FormClosing);
            this.Load += new System.EventHandler(this.ProcessTool_Load);
            ((System.ComponentModel.ISupportInitialize)(this.ribbonControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pnlExcelctrl)).EndInit();
            this.pnlExcelctrl.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.spreadsheetBarController1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevExpress.LookAndFeel.DefaultLookAndFeel defaultLookAndFeel1;
        private DevExpress.XtraEditors.PanelControl pnlExcelctrl;
        private DevExpress.XtraSplashScreen.SplashScreenManager splashScreenManager1;
        private DevExpress.XtraSpreadsheet.UI.CommonRibbonPageGroup commonRibbonPageGroup1;
        private DevExpress.XtraSpreadsheet.UI.CommonRibbonPageGroup commonRibbonPageGroup2;
        private DevExpress.XtraSpreadsheet.UI.CommonRibbonPageGroup commonRibbonPageGroup3;
        private DevExpress.XtraSpreadsheet.UI.SpreadsheetCommandBarButtonItem spreadsheetCommandBarButtonItem1;
        private DevExpress.XtraBars.Ribbon.RibbonControl ribbonControl1;
        private DevExpress.XtraSpreadsheet.UI.SpreadsheetCommandBarButtonItem spreadsheetCommandBarButtonItem12;
        private DevExpress.XtraSpreadsheet.UI.FileRibbonPage fileRibbonPage1;
        private DevExpress.XtraSpreadsheet.UI.CommonRibbonPageGroup commonRibbonPageGroup4;
        private DevExpress.XtraSpreadsheet.UI.SpreadsheetBarController spreadsheetBarController1;
        private DevExpress.XtraSpreadsheet.UI.SpreadsheetCommandBarButtonItem spreadsheetCommandBarButtonItem6;
        private DevExpress.XtraSpreadsheet.UI.SpreadsheetCommandBarButtonItem spreadsheetCommandBarButtonItem5;
        private DevExpress.XtraBars.BarButtonItem saveButton;
        private DevExpress.XtraSpreadsheet.UI.SpreadsheetCommandBarButtonItem spreadsheetCommandBarButtonItem2;
        private DevExpress.XtraSpreadsheet.SpreadsheetControl spreadsheetControl1;
        private DevExpress.XtraBars.BarButtonItem btnHelp;
        private DevExpress.XtraBars.Ribbon.RibbonPageGroup Inforbn;
        private DevExpress.XtraBars.BarButtonItem btnExport;

    }
}