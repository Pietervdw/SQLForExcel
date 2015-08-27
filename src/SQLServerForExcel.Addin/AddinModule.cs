using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using AddinExpress.MSO;
using SQLServerForExcel_Addin.Extensions;
using Excel = Microsoft.Office.Interop.Excel;

namespace SQLServerForExcel_Addin
{
    /// <summary>
    ///   Add-in Express Add-in Module
    /// </summary>
    [GuidAttribute("667BD87A-B284-4256-860F-F09203890926"), ProgId("SQLServerForExcel_Addin.AddinModule")]
    public class AddinModule : AddinExpress.MSO.ADXAddinModule
    {
        public AddinModule()
        {
            Application.EnableVisualStyles();
            InitializeComponent();
            // Please add any initialization code to the AddinInitialize event handler
        }

        private AddinExpress.XL.ADXExcelTaskPanesManager taskPanesManager;
        private AddinExpress.XL.ADXExcelTaskPanesCollectionItem databaseExplorerTaskPaneItem;
        private ADXRibbonTab databaseRibbonTab;
        private ADXRibbonGroup databaseRibbonGroup;
        private ADXRibbonButton sqlforExcelRibbonButton;
        private ImageList imlIcons;
        private ADXExcelAppEvents excelEvents;
        public bool SheetChangeEvent = true;

        #region Component Designer generated code
        /// <summary>
        /// Required by designer
        /// </summary>
        private System.ComponentModel.IContainer components;

        /// <summary>
        /// Required by designer support - do not modify
        /// the following method
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddinModule));
            this.taskPanesManager = new AddinExpress.XL.ADXExcelTaskPanesManager(this.components);
            this.databaseExplorerTaskPaneItem = new AddinExpress.XL.ADXExcelTaskPanesCollectionItem(this.components);
            this.databaseRibbonTab = new AddinExpress.MSO.ADXRibbonTab(this.components);
            this.databaseRibbonGroup = new AddinExpress.MSO.ADXRibbonGroup(this.components);
            this.sqlforExcelRibbonButton = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.imlIcons = new System.Windows.Forms.ImageList(this.components);
            this.excelEvents = new AddinExpress.MSO.ADXExcelAppEvents(this.components);
            // 
            // taskPanesManager
            // 
            this.taskPanesManager.Items.Add(this.databaseExplorerTaskPaneItem);
            this.taskPanesManager.SetOwner(this);
            // 
            // databaseExplorerTaskPaneItem
            // 
            this.databaseExplorerTaskPaneItem.AllowedDropPositions = ((AddinExpress.XL.ADXExcelAllowedDropPositions)((AddinExpress.XL.ADXExcelAllowedDropPositions.Right | AddinExpress.XL.ADXExcelAllowedDropPositions.Left)));
            this.databaseExplorerTaskPaneItem.AlwaysShowHeader = true;
            this.databaseExplorerTaskPaneItem.CloseButton = true;
            this.databaseExplorerTaskPaneItem.IsDragDropAllowed = true;
            this.databaseExplorerTaskPaneItem.Position = AddinExpress.XL.ADXExcelTaskPanePosition.Right;
            this.databaseExplorerTaskPaneItem.TaskPaneClassName = "SQLServerForExcel_Addin.DatabaseExplorerTaskPane";
            this.databaseExplorerTaskPaneItem.UseOfficeThemeForBackground = true;
            // 
            // databaseRibbonTab
            // 
            this.databaseRibbonTab.Caption = "Database";
            this.databaseRibbonTab.Controls.Add(this.databaseRibbonGroup);
            this.databaseRibbonTab.Id = "adxRibbonTab_85a6421e5ca84f33806886691942c8c1";
            this.databaseRibbonTab.IdMso = "TabData";
            this.databaseRibbonTab.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // databaseRibbonGroup
            // 
            this.databaseRibbonGroup.Caption = "Database";
            this.databaseRibbonGroup.Controls.Add(this.sqlforExcelRibbonButton);
            this.databaseRibbonGroup.Id = "adxRibbonGroup_81428551842449ca932d3fa453231758";
            this.databaseRibbonGroup.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.databaseRibbonGroup.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // sqlforExcelRibbonButton
            // 
            this.sqlforExcelRibbonButton.Caption = "SQL Server for Excel";
            this.sqlforExcelRibbonButton.Glyph = global::SQLServerForExcel_Addin.Properties.Resources.SSMS;
            this.sqlforExcelRibbonButton.Id = "adxRibbonButton_c3ef0c24017d4be0b1a084febd77f725";
            this.sqlforExcelRibbonButton.ImageList = this.imlIcons;
            this.sqlforExcelRibbonButton.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.sqlforExcelRibbonButton.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.sqlforExcelRibbonButton.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.sqlforExcelRibbonButton.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.sqlforExcelRibbonButton_OnClick);
            // 
            // imlIcons
            // 
            this.imlIcons.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imlIcons.ImageStream")));
            this.imlIcons.TransparentColor = System.Drawing.Color.Transparent;
            this.imlIcons.Images.SetKeyName(0, "logoico.ico");
            // 
            // excelEvents
            // 
            this.excelEvents.SheetChange += new AddinExpress.MSO.ADXExcelSheet_EventHandler(this.excelEvents_SheetChange);
            // 
            // AddinModule
            // 
            this.AddinName = "SQLServerForExcel_Addin";
            this.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaExcel;

        }
        #endregion

        #region Add-in Express automatic code

        // Required by Add-in Express - do not modify
        // the methods within this region

        public override System.ComponentModel.IContainer GetContainer()
        {
            if (components == null)
                components = new System.ComponentModel.Container();
            return components;
        }

        [ComRegisterFunctionAttribute]
        public static void AddinRegister(Type t)
        {
            AddinExpress.MSO.ADXAddinModule.ADXRegister(t);
        }

        [ComUnregisterFunctionAttribute]
        public static void AddinUnregister(Type t)
        {
            AddinExpress.MSO.ADXAddinModule.ADXUnregister(t);
        }

        public override void UninstallControls()
        {
            base.UninstallControls();
        }

        #endregion

        public static new AddinModule CurrentInstance
        {
            get
            {
                return AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule;
            }
        }

        public Excel._Application ExcelApp
        {
            get
            {
                return (HostApplication as Excel._Application);
            }
        }

        private void excelEvents_SheetChange(object sender, object sheet, object range)
        {
            Excel.Worksheet changedSheet = null;
            Excel.Range changedRange = null;

            try
            {
                changedSheet = sheet as Excel.Worksheet;
                if (SheetChangeEvent && changedSheet.ConnectedToDb())
                {
                    changedRange = range as Excel.Range;
                    changedSheet.AddChangedRow(changedRange);
                }

            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
            finally
            {
                //if (changedSheet != null) Marshal.ReleaseComObject(changedSheet); // Disposed in DatabaseExplorerPane.tvTables_NodeMouseDoubleClick
            }
        }

        private void sqlforExcelRibbonButton_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            databaseExplorerTaskPaneItem.ShowTaskPane();
        }

    }
}

