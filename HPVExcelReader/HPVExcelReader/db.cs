using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Transactions;

namespace HPVExcelReader
{
    public class TempDodMergeCommand
    {
        internal string tempName { get; set; }

        public void DoMerge()
        {
            var dropTableSql = string.Empty;
            var sql = string.Empty;
            //we clear the current database table
            sql = "select top 1 srcfile from {0}";
            var file = db.GetListText(string.Format(sql, tempName)).FirstOrDefault();
            //we delete from the mergeto table any matching records by file
            sql = "delete from StagingTable where srcfile = '{0}'";
            db.ExecSql(string.Format(sql, file));

            //we copy from the temp
            sql = "insert into StagingTable(dirpath,srcfile,FacilityID,IndicatorID,ReferenceYear,ReferenceMonth,Sex,AgeGroup,Number) select dirpath,srcfile,FacilityID,IndicatorID,ReferenceYear,ReferenceMonth,Sex,AgeGroup,Number from {0}";
            db.ExecSql(string.Format(sql, tempName));

            sql = string.Format("drop table {0};", tempName);
            db.ExecSql(sql);
        }
    }

    public class SaveTableToDbCommand : IQueryHelper<IEnumerable<string>>
    {
        public Action<string> Alert { get; set; }
        //public DbHelper Db { get; set; }
        public bool IsWebData { get; set; }

        public IDisplayProgress progressDisplayHelper { get; set; }
        public DataSet TargetDataset { get; internal set; }

        public IEnumerable<string> Execute()
        {
            var tablecount = TargetDataset.Tables.Count;
            for (var i = 0; i < tablecount; i++)
            {
                //we copy to server
                var table = TargetDataset.Tables[i];
                var targetTable = table.TableName;
                //we create the table
                var builder = new StringBuilder();
                var res =
                    string.Format("create table {0} ({1})", targetTable,
                    string.Join(",",
                    (
                from DataColumn dc in table.Columns
                select "[" + dc.ColumnName + "] varchar(256)")));

                //initialise db

                int recordsImported = -5;
                using (var transaction = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(2, 0, 0)))
                {
                    //and create the temp table
                    db.ExecSql(res);

                    //use bulkcopy to write table to db

                    db.WriteTableToDb(table, targetTable);

                    //we check how many records in the table
                    recordsImported = db.GetScalar("select count(*) from " + targetTable);

                    transaction.Complete();
                }
            }
            //MessageBox.Show("Records will be imported " + recordsImported);
            return new List<string>();
        }
    }
    public static class db
    {
        public static string checkLength2(this string str)
        {
            str = str.Trim().Replace(",", "-").Replace("\"", "").Replace("\t", "").Replace("\n", "").Trim();
            return str.Length > 235 ? str.Substring(0, 235) : str;
        }
        public static string checkLength(this string str)
        {
            str = str.Trim().Replace(",", "-").Replace("\"", "").Replace("\t", "").Replace("\n", "").Replace(" ", "");
            return str.Length > 235 ? str.Substring(0, 235) : str;
        }
        private static List<object> GetList(string sqlStatement)
        {
            var res = new List<object>();
            using (var conn = new SqlConnection(_connString))
            using (var cmd = new SqlCommand(sqlStatement) { Connection = conn, CommandTimeout = 360 })
            {
                conn.Open();
                var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var item = reader[0];
                    if (item == DBNull.Value)
                        continue;

                    res.Add(reader[0]);
                }
                conn.Close();
            }
            return res;
        }

        public static Dictionary<object, object> GetLookups(string sqlStatement)
        {
            var res = new Dictionary<object, object>();
            using (var conn = new SqlConnection(_connString))
            using (var cmd = new SqlCommand(sqlStatement) { Connection = conn, CommandTimeout = 360 })
            {
                conn.Open();
                var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var key = reader[0];
                    var value = reader[1];
                    if (key == DBNull.Value || value == DBNull.Value)
                        continue;

                    res.Add(key, value);
                }
                conn.Close();
            }
            return res;
        }
        //
        internal static List<string> GetListText(string sqlStatement)
        {
            var results = GetList(sqlStatement);
            return (from item in results
                    select Convert.ToString(item)).ToList();
        }

        internal static int GetScalar(string sqlString)
        {
            var res = -1;
            using (var conn = new SqlConnection(_connString))
            using (var cmd = new SqlCommand(sqlString) { Connection = conn, CommandTimeout = 360 })
            {
                conn.Open();
                var dbRes = cmd.ExecuteScalar();
                res = Convert.ToInt32(dbRes);
                conn.Close();
            }
            return res;
        }

        internal static void ExecSql(string res)
        {
            using (var conn = new SqlConnection(_connString))
            using (var cmd = new SqlCommand(res) { Connection = conn, CommandTimeout = 360 })
            {
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
            }
        }

        internal static int ExecSql(string sqlString, CommandParam para)
        {
            var res = -1;
            using (var conn = new SqlConnection(_connString))
            using (var cmd = new SqlCommand(sqlString) { Connection = conn, CommandTimeout = 360 })
            {
                foreach (var p in para.Parameters)
                {
                    cmd.Parameters.AddWithValue(p.Name, p.Value);
                }

                conn.Open();
                res = cmd.ExecuteNonQuery();
                conn.Close();
            }
            return res;
        }

        static string _connString = @"Data Source='.\SQLDEV';Initial Catalog='JhpiegoDB_HPV'; Integrated Security=true";
        public static void WriteTableToDb(this DataTable table, string targetTable)
        {
            //use SqlBulkCopy to write to the server
            using (var conn = new SqlConnection(_connString))
            using (var bcp = new SqlBulkCopy(conn) { DestinationTableName = targetTable, BulkCopyTimeout = 360 })
            {
                foreach (DataColumn c in table.Columns)
                {
                    //if()
                    bcp.ColumnMappings.Add(c.ColumnName, c.ColumnName);
                }
                conn.Open();
                table.AcceptChanges();
                //bcp.WriteToServer(table);
                try
                {
                    bcp.WriteToServer(table);
                    //sqlTran.Commit();
                }
                catch (SqlException ex)
                {
                    if (ex.Message.Contains("Received an invalid column length from the bcp client for colid"))
                    {
                        string pattern = @"\d+";
                        Match match = Regex.Match(ex.Message.ToString(), pattern);
                        var index = Convert.ToInt32(match.Value) - 1;

                        FieldInfo fi = typeof(SqlBulkCopy).GetField("_sortedColumnMappings", BindingFlags.NonPublic | BindingFlags.Instance);
                        var sortedColumns = fi.GetValue(bcp);
                        var items = (Object[])sortedColumns.GetType().GetField("_items", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(sortedColumns);

                        FieldInfo itemdata = items[index].GetType().GetField("_metadata", BindingFlags.NonPublic | BindingFlags.Instance);
                        var metadata = itemdata.GetValue(items[index]);

                        var column = metadata.GetType().GetField("column", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance).GetValue(metadata);
                        var length = metadata.GetType().GetField("length", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance).GetValue(metadata);
                        throw new FormatException(String.Format("Column: {0} contains data with a length greater than: {1}", column, length));
                    }

                    throw;
                }
                bcp.Close();
            }
        }
    }
}
