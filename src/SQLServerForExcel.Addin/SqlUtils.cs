using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;

namespace SQLServerForExcel_Addin
{
    public static class SqlUtils
    {
        public static List<string> GetAllSQLServers()
        {
            List<string> returnValue = new List<string>();
            SqlDataSourceEnumerator servers = SqlDataSourceEnumerator.Instance;
            DataTable serversTable = servers.GetDataSources();

            foreach (DataRow row in serversTable.Rows)
            {
                returnValue.Add(String.Format("{0} {1}", row[0], row[1]));
            }
            return returnValue;
        }

        public static List<string> GetAllDatabases(string connectionString)
        {
            List<string> returnValue = new List<string>();
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand("EXEC sp_databases;", conn))
                {
                    SqlDataReader dbReader;
                    conn.Open();
                    dbReader = cmd.ExecuteReader();
                    while (dbReader.Read())
                    {
                        returnValue.Add(dbReader["DATABASE_NAME"].ToString());
                    }
                }
                conn.Close();
            }
            return returnValue;
        }

        public static List<string> GetAllTables(string connectionString)
        {
            List<string> returnValue = new List<string>();
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder() { ConnectionString = connectionString };
            string _dbName = builder.InitialCatalog;
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand(String.Format("EXEC sp_tables @table_name = '%',@table_qualifier = '{0}',@table_type = \"'Table'\";", _dbName), conn))
                {
                    SqlDataReader dbReader;
                    conn.Open();
                    dbReader = cmd.ExecuteReader();
                    while (dbReader.Read())
                    {
                        returnValue.Add(String.Format("{0}.{1}", dbReader["TABLE_OWNER"], dbReader["TABLE_NAME"]));
                    }
                }
                conn.Close();
            }
            return returnValue;
        }

        public static List<string> GetAllColumns(string connectionString, string tableName)
        {
            List<string> returnValue = new List<string>();
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder() { ConnectionString = connectionString };

            tableName = tableName.Replace("dbo.", "");
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand(String.Format("EXEC sp_columns '{0}';", tableName), conn))
                {
                    SqlDataReader dbReader;
                    conn.Open();
                    dbReader = cmd.ExecuteReader();
                    while (dbReader.Read())
                    {
                        returnValue.Add(dbReader["COLUMN_NAME"].ToString());
                    }
                }
                conn.Close();
            }
            return returnValue;
        }

        public static string GetPrimaryKey(string connectionString, string tableName)
        {
            string[] splitString = tableName.Split('.');
            string returnValue = string.Empty;

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand(String.Format("SELECT B.COLUMN_NAME FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS A, INFORMATION_SCHEMA.CONSTRAINT_COLUMN_USAGE B WHERE CONSTRAINT_TYPE = 'PRIMARY KEY' AND A.CONSTRAINT_NAME = B.CONSTRAINT_NAME And A.TABLE_NAME = '{0}'", splitString[1]), conn))
                {
                    conn.Open();
                    returnValue = cmd.ExecuteScalar().ToString();
                }
                conn.Close();
            }
            return returnValue;
        }

    }
}