﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ZambiaDataManager.Storage
{
    public class DbFactory
    {
        public static string InstanceName = string.Empty;
        public static string ServerName = "ZM-VLUS56";
        public static string MachineNameDevDefault = "ZM-9W48GC2";
        public static string MachineNameDevOther = "SUPER-LAP";
        public static string password = "";
        public static string username = "";

        public static ConnectionBuilder GetDefaultConnection(ProjectName projectName, bool getAlternate = false)
        {
            var defaultServerName = ServerName;
            var defaultSqlExpress = InstanceName;
            if (getAlternate)
            {
                defaultServerName = MachineNameDevDefault;
                defaultSqlExpress = "SQL2014";
            }

            ConnectionBuilder connBuilder = null;
            switch (projectName)
            {
                case ProjectName.DOD:
                    {
                        connBuilder = new ConnectionBuilder()
                        {
                            User = username,
                            Password = password,
                            DatabaseName = "JhpiegoDb_DOD",
                            InstanceName = defaultSqlExpress,
                            ServerName = defaultServerName
                        };
                        break;
                    }
                case ProjectName.IHP_VMMC:
                    {
                        connBuilder = new ConnectionBuilder()
                        {
                            User = username,
                            Password = password,
                            DatabaseName = "JhpiegoDb_IhpVmmc",
                            InstanceName = defaultSqlExpress,
                            ServerName = defaultServerName };
                        break;
                    }
                case ProjectName.IHP_Capacity_Building_and_Training:
                    {
                        connBuilder = new ConnectionBuilder()
                        {
                            User = username,
                            Password = password,
                            DatabaseName = "JhpiegoDb_IhpTraining",
                            InstanceName = defaultSqlExpress,
                            ServerName = defaultServerName };
                        break;
                    }
                case ProjectName.MCSP:
                    {
                        connBuilder = new ConnectionBuilder()
                        {
                            User = username,
                            Password = password,
                            DatabaseName = "JhpiegoDb_MCSP",
                            InstanceName = defaultSqlExpress,
                            ServerName = defaultServerName
                        };
                        break;
                    }
                case ProjectName.General:
                    {
                        connBuilder = new ConnectionBuilder()
                        {
                            User = username,
                            Password = password,
                            DatabaseName = "JhpiegoDb_General",
                            InstanceName = defaultSqlExpress,
                            ServerName = defaultServerName };
                        break;
                    }
            }
            return connBuilder;
        }
    }

    public class ConnectionBuilder
    {
        const string connStringX = @"Data Source = {0}\{1}; Initial Catalog = {2}; User = {3}; Password = {4}";
        const string connStringForDefaultInstance = @"Data Source = {0}; Initial Catalog = {1}; User = {2}; Password = {3}";

        //const string connStringX = @"Data Source = {0}\{1}; Initial Catalog = {2}; Integrated Security = true";
        //const string connStringForDefaultInstance = @"Data Source = {0}; Initial Catalog = {1}; Integrated Security = true";

        public string GetConnectionString()
        {
            if (string.IsNullOrWhiteSpace(InstanceName))
            {
                return string.Format(connStringForDefaultInstance, ServerName, DatabaseName, User, Password);
            }
            return string.Format(connStringX, ServerName, InstanceName, DatabaseName, User, Password);
        }
        public string ServerName { get; set; }
        public string InstanceName { get; set; }

        public string DatabaseName { get; set; }
        public string ConnectionString { get; internal set; }

        public string User { get; set; }
        public string Password { get; internal set; }
        //public static string DefaultInstanceName = "default";
        //public bool IntegratedSecurity { get; set; }
        //public string UserName { get; set; }
        //public string Password { get; set; }
    }

    public class DbHelper
    {
        ConnectionBuilder _connBuilder = null;
        //ProjectName _projectName;
        string _connString;
        public DbHelper(ConnectionBuilder builder)
        {
            _connBuilder = builder;
            _connString = _connBuilder.GetConnectionString();
        }

        internal void ExecSql(string res)
        {
            using (var conn = new SqlConnection(_connString))
            using (var cmd = new SqlCommand(res) { Connection = conn, CommandTimeout = 360 })
            {
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
            }
        }

        internal void WriteTableToDb(string targetTable, DataTable table)
        {
            //use SqlBulkCopy to write to the server
            using (var conn = new SqlConnection(_connString))
            using (var bcp = new SqlBulkCopy(conn) { DestinationTableName = targetTable, BulkCopyTimeout = 360 })
            {
                foreach (DataColumn c in table.Columns)
                {
                    bcp.ColumnMappings.Add(c.ColumnName, c.ColumnName);
                }
                conn.Open();
                table.AcceptChanges();
                bcp.WriteToServer(table);
                bcp.Close();
            }
        }
        private List<object> GetList(string sqlStatement)
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

        public Dictionary<object, object> GetLookups(string sqlStatement)
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
        internal List<string> GetListText(string sqlStatement)
        {
            var results = GetList(sqlStatement);
            return (from item in results
                    select Convert.ToString(item)).ToList();
        }

        internal List<int> GetIntList(string sqlStatement)
        {
            var results = GetList(sqlStatement);
            return (from item in results
                    select Convert.ToInt32(item)).ToList();
        }

        internal int ExecSql(string sqlString, CommandParam para)
        {
            var res = -1;
            using (var conn = new SqlConnection(_connString))
            using (var cmd = new SqlCommand(sqlString) { Connection = conn, CommandTimeout=360 })
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

        internal int GetScalar(string sqlString, CommandParam para)
        {
            var res = -1;
            using (var conn = new SqlConnection(_connString))
            using (var cmd = new SqlCommand(sqlString) { Connection = conn, CommandTimeout = 360 })
            {                
                foreach(var p in para.Parameters)
                {
                    cmd.Parameters.AddWithValue(p.Name, p.Value);
                }

                conn.Open();
                var dbRes = cmd.ExecuteScalar();
                res = Convert.ToInt32(dbRes);
                conn.Close();
            }
            return res;
        }

        internal int GetScalar(string sqlString, int timeout)
        {
            var res = -1;
            using (var conn = new SqlConnection(_connString))
            using (var cmd = new SqlCommand(sqlString) { Connection = conn, CommandTimeout=timeout })
            {
                //cmd.Connection.ConnectionTimeout = timeout;
                conn.Open();
                var dbRes = cmd.ExecuteScalar();
                res = Convert.ToInt32(dbRes);
                conn.Close();
            }
            return res;
        }

        internal int GetScalar(string sqlString)
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

        internal DataTable GetTable(string procName, bool isProc,
            List<KeyValuePair<string,object>> parameters)
        {
            var table = new DataTable();
            using (var conn = new SqlConnection(_connString))
            using (var adapter = new SqlDataAdapter(procName, conn))
            {
                adapter.SelectCommand.CommandType = CommandType.StoredProcedure;
                adapter.SelectCommand.CommandTimeout = 360;
                foreach(var par in parameters){
                    adapter.SelectCommand.Parameters.AddWithValue(par.Key, par.Value);
                }
                conn.Open();
                adapter.Fill(table);
                conn.Close();
            }
            return table;
        }
        internal DataTable GetTable(string sqlString)
        {
            var table = new DataTable();
            using (var conn = new SqlConnection(_connString))
            using(var adapter = new SqlDataAdapter(sqlString, conn))
            {
                adapter.SelectCommand.CommandTimeout = 360;
                conn.Open();
                adapter.Fill(table);
                conn.Close();
            }
            return table;
        }

        internal void ExecProc(string procedureName, CommandParam commandParam)
        {
            var table = new DataTable();
            using (var conn = new SqlConnection(_connString))
            using (var cmd = new SqlCommand(procedureName, conn) {
                CommandType = CommandType.StoredProcedure,
                CommandTimeout = 360
            })
            {
                conn.Open();
                foreach(var par in commandParam.Parameters)
                {
                    cmd.Parameters.AddWithValue(par.Name, par.Value);
                }
                cmd.ExecuteNonQuery();
                conn.Close();
            }
        }
    }
}
