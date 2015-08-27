using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml.Linq;
using GenericParsing;
using Microsoft.Data.ConnectionUI;
using Microsoft.Office.Core;
using SQLServerForExcel_Addin.Extensions;
using Excel = Microsoft.Office.Interop.Excel;

namespace SQLServerForExcel_Addin
{
    public partial class DatabaseExplorerTaskPane : AddinExpress.XL.ADXExcelTaskPane
    {
        DataConnectionDialog dcd;
        string dbName = string.Empty;
        string tableName = string.Empty;
        string serverName = string.Empty;
        string connectionString = string.Empty;
        private System.Data.DataTable sourceData = null;

        public Excel._Application ExcelApp
        {
            get
            {
                return (ExcelAppObj as Excel._Application);
            }
        }

        public DatabaseExplorerTaskPane()
        {
            InitializeComponent();
            dcd = new DataConnectionDialog();
        }

        private void btnConnectToDatabase_Click(object sender, EventArgs e)
        {
            DataConnectionConfiguration dcs = new DataConnectionConfiguration(null);
            dcs.LoadConfiguration(dcd);

            if (DataConnectionDialog.Show(dcd) == DialogResult.OK)
            {
                var tables = SqlUtils.GetAllTables(dcd.ConnectionString);

                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
                builder.ConnectionString = dcd.ConnectionString;
                dbName = builder.InitialCatalog;
                serverName = builder.DataSource;
                connectionString = dcd.ConnectionString;

                TreeNode rootNode = new TreeNode(builder.InitialCatalog, 1, 1);
                TreeNode tablesNode = rootNode.Nodes.Add("Tables", "Tables", 2, 2);
                tablesNode.Tag = dcd.ConnectionString;

                foreach (string table in tables)
                {
                    TreeNode tableNode = tablesNode.Nodes.Add(table, table, 3, 3);
                    tableNode.Tag = "table";
                }

                tvTables.Nodes.Add(rootNode);
                tvTables.ExpandAll();
            }
        }

        private void tvTables_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            Excel.Worksheet sheet = null;
            Excel.Range insertionRange = null;
            Excel.QueryTable queryTable = null;
            Excel.QueryTables queryTables = null;
            Excel.Range cellRange = null;
            Excel.CustomProperties sheetProperties = null;
            Excel.CustomProperty primaryKeyProperty = null;

            SqlConnectionStringBuilder builder = null;
            string connString = "OLEDB;Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Data Source=@servername;Initial Catalog=@databasename";
            string connStringSQL = "OLEDB;Provider=SQLOLEDB.1;Persist Security Info=True;User ID=@username;Password=@password;Data Source=@servername;Initial Catalog=@databasename";
            string databaseName = string.Empty;
            string tableName = string.Empty;

            try
            {
                var module = this.AddinModule as AddinModule;
                module.SheetChangeEvent = false;
                tableName = e.Node.Text;
                sheet = ExcelApp.ActiveSheet as Excel.Worksheet;
                cellRange = sheet.Cells;
                insertionRange = cellRange[1, 1] as Excel.Range;
                builder = new SqlConnectionStringBuilder(dcd.ConnectionString);
                databaseName = builder.InitialCatalog;
                if (!builder.IntegratedSecurity)
                    connString = connStringSQL;

                connString =
                    connString.Replace("@servername", builder.DataSource)
                        .Replace("@databasename", databaseName)
                        .Replace("@username", builder.UserID)
                        .Replace("@password", builder.Password);
                queryTables = sheet.QueryTables;

                if (queryTables.Count > 0)
                {
                    queryTable = queryTables.Item(1);
                    queryTable.CommandText = String.Format("SELECT * FROM [{0}].{1}", databaseName, tableName);
                }
                else
                {
                    queryTable = queryTables.Add(connString, insertionRange,
                        String.Format("SELECT * FROM [{0}].{1}", databaseName, tableName));
                }
                queryTable.RefreshStyle = Excel.XlCellInsertionMode.xlOverwriteCells;
                queryTable.PreserveColumnInfo = true;
                queryTable.PreserveFormatting = true;
                queryTable.Refresh(false);

                var primaryKey = SqlUtils.GetPrimaryKey(dcd.ConnectionString, tableName);
                sheet.Name = tableName;
                chPrimaryKey.Text = primaryKey;

                sheetProperties = sheet.CustomProperties;
                primaryKeyProperty = sheetProperties.Add("PrimaryKey", primaryKey);
                module.SheetChangeEvent = true;
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
            finally
            {
                if (primaryKeyProperty != null) Marshal.ReleaseComObject(primaryKeyProperty);
                if (sheetProperties != null) Marshal.ReleaseComObject(sheetProperties);
                if (cellRange != null) Marshal.ReleaseComObject(cellRange);
                if (queryTables != null) Marshal.ReleaseComObject(queryTables);
                if (queryTable != null) Marshal.ReleaseComObject(queryTable);
                if (insertionRange != null) Marshal.ReleaseComObject(insertionRange);
                if (sheet != null) Marshal.ReleaseComObject(sheet);
            }
        }

        private void tabPageSheetChanges_Enter(object sender, EventArgs e)
        {
            //RefreshChanges();
        }

        public void RefreshChanges()
        {
            Excel.Worksheet activeSheet = null;
            Excel.CustomProperty changesProperty = null;
            string xml = string.Empty;

            try
            {
                activeSheet = ExcelApp.ActiveSheet as Excel.Worksheet;
                changesProperty = activeSheet.GetProperty("UncommittedChanges");
                lvSheetChanges.Items.Clear();
                if (changesProperty != null)
                {
                    lvSheetChanges.Visible = true;
                    xml = ToSafeXml("<uncommittedchanges>" + changesProperty.Value.ToString() + "</uncommittedchanges>");
                    XDocument doc = XDocument.Parse(xml);
                    foreach (var dm in doc.Descendants("row"))
                    {
                        ListViewItem item = new ListViewItem(new string[]
                        {
                            dm.Attribute("key").Value,
                            dm.Attribute("column").Value,
                            dm.Value
                        });
                        lvSheetChanges.Items.Add(item);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
            finally
            {
                //if (activeSheet != null) Marshal.ReleaseComObject(activeSheet);
                //if (changesProperty != null) Marshal.ReleaseComObject(changesProperty);
            }
        }

        private void btnBrowseForDataFile_Click(object sender, EventArgs e)
        {
            diagOpenFile.ShowDialog();
            if (!String.IsNullOrEmpty(diagOpenFile.FileName))
            {
                txtDataFile.Text = diagOpenFile.FileName;

                using (GenericParserAdapter parser = new GenericParserAdapter(diagOpenFile.FileName))
                {
                    parser.FirstRowHasHeader = true;
                    parser.Read();
                    sourceData = parser.GetDataTable();
                }

                foreach (DataColumn column in sourceData.Columns)
                {
                    cboColumnNames.Items.Add(column.ColumnName);
                }
            }
        }

        private void btnInsertDataToSelection_Click(object sender, EventArgs e)
        {
            Excel.Range selectedRange = null;
            int cellCount = 0;

            try
            {
                selectedRange = ExcelApp.Selection as Excel.Range;
                cellCount = selectedRange.Count;
                if (selectedRange != null)
                {
                    List<DataRow> randomData = sourceData.Rows.OfType<DataRow>().Shuffle(new Random()).Take(cellCount).ToList();
                    int numCount = 0;
                    foreach (Excel.Range cell in selectedRange.Cells)
                    {
                        cell.Value = randomData[numCount][cboColumnNames.Text];
                        numCount++;
                    }
                }
            }
            finally
            {
                if (selectedRange != null) Marshal.ReleaseComObject(selectedRange);
            }
        }

        private void lblRefresh_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            RefreshChanges();
        }

        public static string ToSafeXml(string xmlString)
        {
            try
            {
                if ((xmlString != null))
                {
                    xmlString = xmlString.Replace("&", "&amp;");
                    xmlString = xmlString.Replace("'", "''");
                    //xmlString = xmlString.Replace(">", "&gt;");
                    //xmlString = xmlString.Replace("<", "&lt;");
                    //xmlString = xmlString.Replace("\"", "&quot;");
                    xmlString = xmlString.Replace("–", "-");
                    return xmlString;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception Er)
            {
                return "";
            }
        }

        private void btnSaveChangesToFile_Click(object sender, EventArgs e)
        {
            Excel.Worksheet sheet = null;
            Excel.CustomProperty primaryKeyProperty = null;
            string primaryKey = string.Empty;
            string tableName = string.Empty;

            try
            {
                sheet = ExcelApp.ActiveSheet as Excel.Worksheet;
                if (sheet != null)
                {
                    tableName = sheet.Name;
                    primaryKeyProperty = sheet.GetProperty("PrimaryKey");
                    if (primaryKeyProperty != null)
                    {
                        primaryKey = primaryKeyProperty.Value.ToString();
                        string sql = sheet.ChangesToSql(tableName, primaryKey);

                        diagSaveFile.ShowDialog();
                        if (!string.IsNullOrEmpty(diagSaveFile.FileName))
                        {
                            File.WriteAllText(diagSaveFile.FileName, sql);
                            RefreshSheetData();
                        }
                    }
                }
            }
            finally
            {
                if (primaryKeyProperty != null) Marshal.ReleaseComObject(primaryKeyProperty);
                if (sheet != null) Marshal.ReleaseComObject(sheet);
            }
        }

        private void RefreshSheetData()
        {
            Excel.Worksheet sheet = null;
            Excel.QueryTables queryTables = null;
            Excel.QueryTable queryTable = null;
            Excel.CustomProperty changesProperty = null;

            try
            {
                var module = this.AddinModule as AddinModule;
                module.SheetChangeEvent = false;
                sheet = ExcelApp.ActiveSheet as Excel.Worksheet;
                if (sheet != null)
                {
                    queryTables = sheet.QueryTables;

                    if (queryTables.Count > 0)
                    {
                        queryTable = queryTables.Item(1);
                        queryTable.RefreshStyle = Excel.XlCellInsertionMode.xlOverwriteCells;
                        queryTable.PreserveColumnInfo = true;
                        queryTable.PreserveFormatting = true;
                        queryTable.Refresh(false);
                    }
                    changesProperty = sheet.GetProperty("uncommittedchanges");
                    if (changesProperty != null)
                        changesProperty.Delete();
                }
                module.SheetChangeEvent = true;
            }
            finally
            {
                if (changesProperty != null) Marshal.ReleaseComObject(changesProperty);
                if (queryTable != null) Marshal.ReleaseComObject(queryTable);
                if (queryTables != null) Marshal.ReleaseComObject(queryTables);
                if (sheet != null) Marshal.ReleaseComObject(sheet);
            }
        }

        private void btnRefreshData_Click(object sender, EventArgs e)
        {
            RefreshSheetData();
        }

        private void btnApplyChangesToDb_Click(object sender, EventArgs e)
        {

            Excel.Worksheet sheet = null;
            Excel.CustomProperty primaryKeyProperty = null;
            string primaryKey = string.Empty;
            string tableName = string.Empty;

            try
            {
                if (MessageBox.Show("This will commit the changes to the database. This action cannot be reversed. Are you sure?", "Confirm", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    sheet = ExcelApp.ActiveSheet as Excel.Worksheet;
                    if (sheet != null)
                    {
                        tableName = sheet.Name;
                        primaryKeyProperty = sheet.GetProperty("PrimaryKey");
                        if (primaryKeyProperty != null)
                        {
                            primaryKey = primaryKeyProperty.Value.ToString();
                            string sql = sheet.ChangesToSql(tableName, primaryKey);

                            if (!string.IsNullOrEmpty(sql))
                            {
                                using (SqlConnection conn = new SqlConnection(dcd.ConnectionString))
                                {
                                    SqlCommand cmd = new SqlCommand(sql, conn);
                                    if (conn.State == ConnectionState.Closed)
                                    {
                                        conn.Open();
                                        cmd.ExecuteNonQuery();
                                    }
                                }
                                RefreshSheetData();
                            }
                        }
                    }
                }
            }
            finally
            {
                if (primaryKeyProperty != null) Marshal.ReleaseComObject(primaryKeyProperty);
                if (sheet != null) Marshal.ReleaseComObject(sheet);
            }

        }

        //http://www.codeproject.com/Articles/11698/A-Portable-and-Efficient-Generic-Parser-for-Flat-F
    }
}
