namespace SQLServerForExcel_Addin
{
    partial class DatabaseExplorerTaskPane
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
  
        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }
  
        #region Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DatabaseExplorerTaskPane));
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.btnConnectToDatabase = new System.Windows.Forms.ToolStripButton();
            this.btnRefreshData = new System.Windows.Forms.ToolStripButton();
            this.btnApplyChangesToDb = new System.Windows.Forms.ToolStripButton();
            this.btnSaveChangesToFile = new System.Windows.Forms.ToolStripButton();
            this.imlIcons = new System.Windows.Forms.ImageList(this.components);
            this.tvTables = new System.Windows.Forms.TreeView();
            this.tabMain = new System.Windows.Forms.TabControl();
            this.tabDatabaseExplorer = new System.Windows.Forms.TabPage();
            this.tabPageSheetChanges = new System.Windows.Forms.TabPage();
            this.lblRefresh = new System.Windows.Forms.LinkLabel();
            this.lvSheetChanges = new System.Windows.Forms.ListView();
            this.chPrimaryKey = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.chColName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.chNewValue = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.tabPageDataGeneration = new System.Windows.Forms.TabPage();
            this.btnInsertDataToSelection = new System.Windows.Forms.Button();
            this.cboColumnNames = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnBrowseForDataFile = new System.Windows.Forms.Button();
            this.txtDataFile = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.diagOpenFile = new System.Windows.Forms.OpenFileDialog();
            this.diagSaveFile = new System.Windows.Forms.SaveFileDialog();
            this.toolStrip1.SuspendLayout();
            this.tabMain.SuspendLayout();
            this.tabDatabaseExplorer.SuspendLayout();
            this.tabPageSheetChanges.SuspendLayout();
            this.tabPageDataGeneration.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.BackColor = System.Drawing.Color.Transparent;
            this.toolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnConnectToDatabase,
            this.btnRefreshData,
            this.btnApplyChangesToDb,
            this.btnSaveChangesToFile});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.System;
            this.toolStrip1.Size = new System.Drawing.Size(322, 25);
            this.toolStrip1.TabIndex = 0;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // btnConnectToDatabase
            // 
            this.btnConnectToDatabase.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnConnectToDatabase.Image = ((System.Drawing.Image)(resources.GetObject("btnConnectToDatabase.Image")));
            this.btnConnectToDatabase.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnConnectToDatabase.Name = "btnConnectToDatabase";
            this.btnConnectToDatabase.Size = new System.Drawing.Size(23, 22);
            this.btnConnectToDatabase.Text = "Connect to Database";
            this.btnConnectToDatabase.Click += new System.EventHandler(this.btnConnectToDatabase_Click);
            // 
            // btnRefreshData
            // 
            this.btnRefreshData.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnRefreshData.Image = ((System.Drawing.Image)(resources.GetObject("btnRefreshData.Image")));
            this.btnRefreshData.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnRefreshData.Name = "btnRefreshData";
            this.btnRefreshData.Size = new System.Drawing.Size(23, 22);
            this.btnRefreshData.Text = "Refresh data (Will revert all changes)";
            this.btnRefreshData.Click += new System.EventHandler(this.btnRefreshData_Click);
            // 
            // btnApplyChangesToDb
            // 
            this.btnApplyChangesToDb.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnApplyChangesToDb.Image = ((System.Drawing.Image)(resources.GetObject("btnApplyChangesToDb.Image")));
            this.btnApplyChangesToDb.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnApplyChangesToDb.Name = "btnApplyChangesToDb";
            this.btnApplyChangesToDb.Size = new System.Drawing.Size(23, 22);
            this.btnApplyChangesToDb.Text = "Apply changes to Table";
            this.btnApplyChangesToDb.Click += new System.EventHandler(this.btnApplyChangesToDb_Click);
            // 
            // btnSaveChangesToFile
            // 
            this.btnSaveChangesToFile.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnSaveChangesToFile.Image = ((System.Drawing.Image)(resources.GetObject("btnSaveChangesToFile.Image")));
            this.btnSaveChangesToFile.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnSaveChangesToFile.Name = "btnSaveChangesToFile";
            this.btnSaveChangesToFile.Size = new System.Drawing.Size(23, 22);
            this.btnSaveChangesToFile.Text = "Save Changes To File";
            this.btnSaveChangesToFile.Click += new System.EventHandler(this.btnSaveChangesToFile_Click);
            // 
            // imlIcons
            // 
            this.imlIcons.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imlIcons.ImageStream")));
            this.imlIcons.TransparentColor = System.Drawing.Color.Magenta;
            this.imlIcons.Images.SetKeyName(0, "database_connect");
            this.imlIcons.Images.SetKeyName(1, "database.bmp");
            this.imlIcons.Images.SetKeyName(2, "Folder.ico");
            this.imlIcons.Images.SetKeyName(3, "TableHS.png");
            this.imlIcons.Images.SetKeyName(4, "Run.bmp");
            // 
            // tvTables
            // 
            this.tvTables.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tvTables.ImageIndex = 0;
            this.tvTables.ImageList = this.imlIcons;
            this.tvTables.Location = new System.Drawing.Point(3, 3);
            this.tvTables.Name = "tvTables";
            this.tvTables.SelectedImageIndex = 0;
            this.tvTables.Size = new System.Drawing.Size(308, 392);
            this.tvTables.TabIndex = 1;
            this.tvTables.NodeMouseDoubleClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.tvTables_NodeMouseDoubleClick);
            // 
            // tabMain
            // 
            this.tabMain.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabMain.Controls.Add(this.tabDatabaseExplorer);
            this.tabMain.Controls.Add(this.tabPageSheetChanges);
            this.tabMain.Controls.Add(this.tabPageDataGeneration);
            this.tabMain.Location = new System.Drawing.Point(0, 28);
            this.tabMain.Name = "tabMain";
            this.tabMain.SelectedIndex = 0;
            this.tabMain.Size = new System.Drawing.Size(322, 424);
            this.tabMain.TabIndex = 2;
            // 
            // tabDatabaseExplorer
            // 
            this.tabDatabaseExplorer.Controls.Add(this.tvTables);
            this.tabDatabaseExplorer.Location = new System.Drawing.Point(4, 22);
            this.tabDatabaseExplorer.Name = "tabDatabaseExplorer";
            this.tabDatabaseExplorer.Padding = new System.Windows.Forms.Padding(3);
            this.tabDatabaseExplorer.Size = new System.Drawing.Size(314, 398);
            this.tabDatabaseExplorer.TabIndex = 0;
            this.tabDatabaseExplorer.Text = "Database Explorer";
            this.tabDatabaseExplorer.UseVisualStyleBackColor = true;
            // 
            // tabPageSheetChanges
            // 
            this.tabPageSheetChanges.Controls.Add(this.lblRefresh);
            this.tabPageSheetChanges.Controls.Add(this.lvSheetChanges);
            this.tabPageSheetChanges.Location = new System.Drawing.Point(4, 22);
            this.tabPageSheetChanges.Name = "tabPageSheetChanges";
            this.tabPageSheetChanges.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageSheetChanges.Size = new System.Drawing.Size(314, 398);
            this.tabPageSheetChanges.TabIndex = 1;
            this.tabPageSheetChanges.Text = "Sheet Changes";
            this.tabPageSheetChanges.UseVisualStyleBackColor = true;
            this.tabPageSheetChanges.Enter += new System.EventHandler(this.tabPageSheetChanges_Enter);
            // 
            // lblRefresh
            // 
            this.lblRefresh.AutoSize = true;
            this.lblRefresh.Image = ((System.Drawing.Image)(resources.GetObject("lblRefresh.Image")));
            this.lblRefresh.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblRefresh.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.lblRefresh.Location = new System.Drawing.Point(9, 7);
            this.lblRefresh.Name = "lblRefresh";
            this.lblRefresh.Size = new System.Drawing.Size(78, 13);
            this.lblRefresh.TabIndex = 1;
            this.lblRefresh.TabStop = true;
            this.lblRefresh.Text = "     Refresh List";
            this.lblRefresh.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lblRefresh_LinkClicked);
            // 
            // lvSheetChanges
            // 
            this.lvSheetChanges.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lvSheetChanges.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.chPrimaryKey,
            this.chColName,
            this.chNewValue});
            this.lvSheetChanges.Location = new System.Drawing.Point(3, 31);
            this.lvSheetChanges.Name = "lvSheetChanges";
            this.lvSheetChanges.Size = new System.Drawing.Size(308, 364);
            this.lvSheetChanges.TabIndex = 0;
            this.lvSheetChanges.UseCompatibleStateImageBehavior = false;
            this.lvSheetChanges.View = System.Windows.Forms.View.Details;
            this.lvSheetChanges.Visible = false;
            // 
            // chPrimaryKey
            // 
            this.chPrimaryKey.Text = "Primary Key";
            this.chPrimaryKey.Width = 80;
            // 
            // chColName
            // 
            this.chColName.Text = "Column Name";
            this.chColName.Width = 100;
            // 
            // chNewValue
            // 
            this.chNewValue.Text = "New Value";
            this.chNewValue.Width = 150;
            // 
            // tabPageDataGeneration
            // 
            this.tabPageDataGeneration.Controls.Add(this.btnInsertDataToSelection);
            this.tabPageDataGeneration.Controls.Add(this.cboColumnNames);
            this.tabPageDataGeneration.Controls.Add(this.label2);
            this.tabPageDataGeneration.Controls.Add(this.btnBrowseForDataFile);
            this.tabPageDataGeneration.Controls.Add(this.txtDataFile);
            this.tabPageDataGeneration.Controls.Add(this.label1);
            this.tabPageDataGeneration.Location = new System.Drawing.Point(4, 22);
            this.tabPageDataGeneration.Name = "tabPageDataGeneration";
            this.tabPageDataGeneration.Size = new System.Drawing.Size(314, 398);
            this.tabPageDataGeneration.TabIndex = 2;
            this.tabPageDataGeneration.Text = "Data Generation";
            this.tabPageDataGeneration.UseVisualStyleBackColor = true;
            // 
            // btnInsertDataToSelection
            // 
            this.btnInsertDataToSelection.Location = new System.Drawing.Point(87, 63);
            this.btnInsertDataToSelection.Name = "btnInsertDataToSelection";
            this.btnInsertDataToSelection.Size = new System.Drawing.Size(169, 23);
            this.btnInsertDataToSelection.TabIndex = 6;
            this.btnInsertDataToSelection.Text = "Insert random data in selection";
            this.btnInsertDataToSelection.UseVisualStyleBackColor = true;
            this.btnInsertDataToSelection.Click += new System.EventHandler(this.btnInsertDataToSelection_Click);
            // 
            // cboColumnNames
            // 
            this.cboColumnNames.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboColumnNames.FormattingEnabled = true;
            this.cboColumnNames.Location = new System.Drawing.Point(87, 35);
            this.cboColumnNames.Name = "cboColumnNames";
            this.cboColumnNames.Size = new System.Drawing.Size(182, 21);
            this.cboColumnNames.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(8, 38);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(71, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Data Column:";
            // 
            // btnBrowseForDataFile
            // 
            this.btnBrowseForDataFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBrowseForDataFile.Location = new System.Drawing.Point(274, 7);
            this.btnBrowseForDataFile.Name = "btnBrowseForDataFile";
            this.btnBrowseForDataFile.Size = new System.Drawing.Size(24, 22);
            this.btnBrowseForDataFile.TabIndex = 2;
            this.btnBrowseForDataFile.Text = "...";
            this.btnBrowseForDataFile.UseVisualStyleBackColor = true;
            this.btnBrowseForDataFile.Click += new System.EventHandler(this.btnBrowseForDataFile_Click);
            // 
            // txtDataFile
            // 
            this.txtDataFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtDataFile.Location = new System.Drawing.Point(87, 8);
            this.txtDataFile.Name = "txtDataFile";
            this.txtDataFile.Size = new System.Drawing.Size(182, 20);
            this.txtDataFile.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(52, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Data file :";
            // 
            // diagOpenFile
            // 
            this.diagOpenFile.Filter = "CSV files|*.csv|All files|*.*";
            this.diagOpenFile.Title = "Select data file";
            // 
            // diagSaveFile
            // 
            this.diagSaveFile.DefaultExt = "sql";
            this.diagSaveFile.Filter = "SQL Files (*.sql)|*.sql|All files|*.*|Text files|*.txt";
            this.diagSaveFile.Title = "Save File As";
            // 
            // DatabaseExplorerTaskPane
            // 
            this.ClientSize = new System.Drawing.Size(322, 464);
            this.Controls.Add(this.tabMain);
            this.Controls.Add(this.toolStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "DatabaseExplorerTaskPane";
            this.Text = "SQL Server for Excel (BETA)";
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.tabMain.ResumeLayout(false);
            this.tabDatabaseExplorer.ResumeLayout(false);
            this.tabPageSheetChanges.ResumeLayout(false);
            this.tabPageSheetChanges.PerformLayout();
            this.tabPageDataGeneration.ResumeLayout(false);
            this.tabPageDataGeneration.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ImageList imlIcons;
        private System.Windows.Forms.ToolStripButton btnConnectToDatabase;
        private System.Windows.Forms.TreeView tvTables;
        private System.Windows.Forms.TabControl tabMain;
        private System.Windows.Forms.TabPage tabDatabaseExplorer;
        private System.Windows.Forms.TabPage tabPageSheetChanges;
        private System.Windows.Forms.TabPage tabPageDataGeneration;
        private System.Windows.Forms.ListView lvSheetChanges;
        private System.Windows.Forms.ColumnHeader chPrimaryKey;
        private System.Windows.Forms.ColumnHeader chColName;
        private System.Windows.Forms.ColumnHeader chNewValue;
        private System.Windows.Forms.ToolStripButton btnApplyChangesToDb;
        private System.Windows.Forms.ToolStripButton btnSaveChangesToFile;
        private System.Windows.Forms.Button btnBrowseForDataFile;
        private System.Windows.Forms.TextBox txtDataFile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.OpenFileDialog diagOpenFile;
        private System.Windows.Forms.ComboBox cboColumnNames;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnInsertDataToSelection;
        private System.Windows.Forms.LinkLabel lblRefresh;
        private System.Windows.Forms.ToolStripButton btnRefreshData;
        private System.Windows.Forms.SaveFileDialog diagSaveFile;
    }
}
