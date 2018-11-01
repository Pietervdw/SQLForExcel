using System;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using AddinExpress.MSO;
using Excel = Microsoft.Office.Interop.Excel;
namespace SQLServerForExcel_Addin.Extensions
{
    public static class WorksheetExtensions
    {
        /// <summary>
        /// Checks whether the sheet has a primary key custom property and
        /// then return true or false indicating it is "connected" to a db table
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns>bool</returns>
        public static bool ConnectedToDb(this Excel.Worksheet sheet)
        {
            Excel.CustomProperties customProperties = null;
            Excel.CustomProperty primaryKeyProperty = null;

            try
            {
                customProperties = sheet.CustomProperties;
                for (int i = 1; i <= customProperties.Count; i++)
                {
                    primaryKeyProperty = customProperties[i];
                    if (primaryKeyProperty != null) Marshal.ReleaseComObject(primaryKeyProperty);
                }
                if (primaryKeyProperty != null)
                    return true;
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
            finally
            {
                if (primaryKeyProperty != null) Marshal.ReleaseComObject(primaryKeyProperty);
                if (customProperties != null) Marshal.ReleaseComObject(customProperties);
            }
            return false;
        }

        public static string PrimaryKey(this Excel.Worksheet sheet)
        {
            Excel.CustomProperties customProperties = null;
            Excel.CustomProperty primaryKeyProperty = null;
            string keyName = null;

            try
            {
                customProperties = sheet.CustomProperties;
                for (int i = 1; i <= customProperties.Count; i++)
                {
                    primaryKeyProperty = customProperties[i];
                    if (primaryKeyProperty.Name == "PrimaryKey")
                    {
                        keyName = primaryKeyProperty.Value.ToString();
                    }
                    if (primaryKeyProperty != null) Marshal.ReleaseComObject(primaryKeyProperty);
                }

            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
            finally
            {
                if (primaryKeyProperty != null) Marshal.ReleaseComObject(primaryKeyProperty);
                if (customProperties != null) Marshal.ReleaseComObject(customProperties);
            }
            return keyName;
        }

        public static string ColumnName(this Excel.Worksheet sheet, int col)
        {
            string columnName = string.Empty;
            Excel.Range columnRange = null;

            try
            {
                string colLetter = ColumnIndexToColumnLetter(col);
                columnRange = sheet.Range[colLetter + "1:" + colLetter + "1"];
                if (columnRange != null)
                {
                    columnName = columnRange.Value.ToString();
                }
            }
            finally
            {
                if (columnRange != null) Marshal.ReleaseComObject(columnRange);
            }
            return columnName;
        }

        public static void AddChangedRow(this Excel.Worksheet sheet, int col, int row)
        {
            Excel.Range columnRange = null;
            Excel.Range primaryKeyColumnRange = null;
            Excel.Range primaryKeyValueRange = null;
            Excel.Range rowValueRange = null;
            Excel.Range sheetCellRange = null;
            Excel.CustomProperty uncommittedChangesProperty = null;
            string primaryKey = string.Empty;
            string primaryKeyDataType = string.Empty;
            object primaryKeyValue = string.Empty;
            string columnName = string.Empty;
            object rowValue = string.Empty;
            string rowValueDataType = string.Empty;

            try
            {
                primaryKey = sheet.PrimaryKey();
                columnRange = sheet.Range["A1:CV1"];
                sheetCellRange = sheet.Cells;
                rowValueRange = sheetCellRange[row, col] as Excel.Range;
                primaryKeyColumnRange = columnRange.Find(primaryKey);

                if (primaryKeyColumnRange != null)
                {
                    primaryKeyValueRange = sheetCellRange[row, primaryKeyColumnRange.Column] as Excel.Range;
                    if (primaryKeyValueRange != null)
                    {
                        primaryKeyValue = primaryKeyValueRange.Value;
                        primaryKeyDataType = primaryKeyValue.GetType().ToString();
                    }
                }

                columnName = sheet.ColumnName(col);
                if (rowValueRange != null)
                {
                    rowValue = rowValueRange.Value;
                    rowValueDataType = rowValue.GetType().ToString();
                }

                string xmlString = "<row key=\"" + primaryKeyValue.ToString() + "\" ";
                xmlString += "keydatatype=\"" + primaryKeyDataType + "\" ";
                xmlString += "column=\"" + columnName + "\" ";
                xmlString += "columndatatype=\"" + rowValueDataType + "\">";
                xmlString += rowValue.ToString();
                xmlString += "</row>";
                xmlString = stripNonValidXMLCharacters(xmlString);

                uncommittedChangesProperty = sheet.GetProperty("UncommittedChanges");
                if (uncommittedChangesProperty == null)
                {
                    uncommittedChangesProperty = sheet.AddProperty("UncommittedChanges", xmlString);
                }
                else
                {
                    uncommittedChangesProperty.Value = uncommittedChangesProperty.Value + xmlString;
                }
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
            finally
            {

            }
        }

        public static void AddChangedRow(this Excel.Worksheet sheet, Excel.Range changedRange)
        {
            Excel.Range columnRange = null;
            Excel.Range primaryKeyColumnRange = null;
            Excel.Range primaryKeyValueRange = null;
            Excel.Range rowValueRange = null;
            Excel.Range sheetCellsRange = null;
            Excel.Range rowsRange = null;
            Excel.Range colsRange = null;
            Excel.CustomProperty uncommittedChangesProperty = null;
            object rowValue = string.Empty;
            string rowValueDataType = string.Empty;
            string primaryKey = string.Empty;
            string primaryKeyDataType = string.Empty;
            object primaryKeyValue = string.Empty;
            string columnName = string.Empty;
            string xmlString = string.Empty;

            try
            {
                primaryKey = sheet.PrimaryKey();
                columnRange = sheet.Range["A1:CV1"];
                sheetCellsRange = sheet.Cells;
                primaryKeyColumnRange = columnRange.Find(primaryKey, LookAt: Excel.XlLookAt.xlWhole);

                rowsRange = changedRange.Rows;
                colsRange = rowsRange.Columns;
                foreach (Excel.Range row in rowsRange)
                {
                    if (primaryKeyColumnRange != null)
                    {
                        int rowNum = row.Row;
                        int colNum = primaryKeyColumnRange.Column;
                        primaryKeyValueRange = sheetCellsRange[rowNum, colNum] as Excel.Range;

                        if (primaryKeyValueRange != null)
                        {
                            primaryKeyValue = primaryKeyValueRange.Value;
                            primaryKeyDataType = primaryKeyValue.GetType().ToString();

                            foreach (Excel.Range col in colsRange)
                            {
                                colNum = col.Column;
                                columnName = sheet.ColumnName(colNum);
                                rowValueRange = sheetCellsRange[rowNum, col.Column] as Excel.Range;
                                if (rowValueRange != null)
                                {
                                    rowValue = rowValueRange.Value;
                                    rowValueDataType = rowValue.GetType().ToString();

                                    xmlString += "<row key=\"" + primaryKeyValue.ToString() + "\" ";
                                    xmlString += "keydatatype=\"" + primaryKeyDataType + "\" ";
                                    xmlString += "column=\"" + columnName + "\" ";
                                    xmlString += "columndatatype=\"" + rowValueDataType + "\">";
                                    xmlString += rowValue.ToString();
                                    xmlString += "</row>";


                                }
                            }
                        }
                    }
                }

                uncommittedChangesProperty = sheet.GetProperty("UncommittedChanges");
                if (uncommittedChangesProperty == null)
                {
                    uncommittedChangesProperty = sheet.AddProperty("UncommittedChanges", xmlString);
                }
                else
                {
                    uncommittedChangesProperty.Value = uncommittedChangesProperty.Value + xmlString;
                }
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
            finally
            {
                if (uncommittedChangesProperty != null) Marshal.ReleaseComObject(uncommittedChangesProperty);
                if (colsRange != null) Marshal.ReleaseComObject(colsRange);
                if (rowsRange != null) Marshal.ReleaseComObject(rowsRange);
                if (sheetCellsRange != null) Marshal.ReleaseComObject(sheetCellsRange);
                if (rowValueRange != null) Marshal.ReleaseComObject(rowValueRange);
                if (primaryKeyValueRange != null) Marshal.ReleaseComObject(primaryKeyValueRange);
                if (primaryKeyColumnRange != null) Marshal.ReleaseComObject(primaryKeyColumnRange);
            }
        }

        public static Excel.CustomProperty AddProperty(this Excel.Worksheet sheet, string propertyName, object propertyValue)
        {
            Excel.CustomProperties customProperties = null;
            Excel.CustomProperty customProperty = null;

            try
            {
                customProperties = sheet.CustomProperties;
                customProperty = customProperties.Add(propertyName, propertyValue);
            }
            finally
            {
                if (customProperties != null) Marshal.ReleaseComObject(customProperties);
            }
            return customProperty;
        }

        public static Excel.CustomProperty GetProperty(this Excel.Worksheet sheet, string propertyName)
        {
            Excel.CustomProperty customProperty = null;
            Excel.CustomProperties customProperties = null;
            try
            {
                customProperties = sheet.CustomProperties;
                for (int i = 1; i <= customProperties.Count; i++)
                {
                    customProperty = customProperties[i];
                    if (customProperty != null && customProperty.Name.ToLower() == propertyName.ToLower())
                    {
                        return customProperty;
                    }
                    else
                    {
                        customProperty = null;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
            finally
            {
                if (customProperties != null) Marshal.ReleaseComObject(customProperties);
            }
            return customProperty;
        }

        public static string ChangesToSql(this Excel.Worksheet sheet, string tableName, string primaryKeyName)
        {
            Excel.CustomProperty customProperty = null;
            string xml = string.Empty;
            string sql = string.Empty;

            try
            {
                customProperty = sheet.GetProperty("uncommittedchanges");
                if (customProperty != null)
                {
                    xml = ToSafeXml("<uncommittedchanges>" + customProperty.Value.ToString() + "</uncommittedchanges>");
                    XDocument doc = XDocument.Parse(xml);
                    foreach (var dm in doc.Descendants("row"))
                    {

                        string key = dm.Attribute("key").Value;
                        string keyDataType = dm.Attribute("keydatatype").Value;
                        string column = dm.Attribute("column").Value;
                        string columnDataType = dm.Attribute("columndatatype").Value;
                        string value = dm.Value;
                      
                        sql += "UPDATE " + tableName + " SET " + column + " = ";
                        
                        if (columnDataType.ToLower().Contains("date") || columnDataType.ToLower().Contains("string") || columnDataType.ToLower().Contains("boolean"))
                        {
                            sql += "'" + value + "'";
                        }
                        else
                        {
                            sql += value;
                        }

                        sql += " WHERE " + primaryKeyName + " = ";

                        if (keyDataType.ToLower().Contains("date") || keyDataType.ToLower().Contains("string"))
                        {
                            sql += "'" + key + "'";
                        }
                        else
                        {
                            sql += key;
                        }

                        sql += Environment.NewLine;
                    }
                }
            }
            finally
            {
                if (customProperty != null) Marshal.ReleaseComObject(customProperty);
            }
            return sql;
        }

        private static string ToSafeXml(string xmlString)
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

        private static string ColumnIndexToColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter;
        }

        private static String stripNonValidXMLCharacters(string textIn)
        {
            StringBuilder textOut = new StringBuilder(); // Used to hold the output.
            char current; // Used to reference the current character.


            if (textIn == null || textIn == string.Empty) return string.Empty; // vacancy test.
            for (int i = 0; i < textIn.Length; i++)
            {
                current = textIn[i];


                if ((current == 0x9 || current == 0xA || current == 0xD) ||
                    ((current >= 0x20) && (current <= 0xD7FF)) ||
                    ((current >= 0xE000) && (current <= 0xFFFD)) ||
                    ((current >= 0x10000) && (current <= 0x10FFFF)))
                {
                    textOut.Append(current);
                }
            }
            return textOut.ToString();
        }
    }
}
