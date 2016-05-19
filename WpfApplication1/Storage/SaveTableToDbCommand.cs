using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;

namespace ZambiaDataManager.Storage
{
    public class SaveTableToDbCommand : IQueryHelper<IEnumerable<string>>
    {
        public Action<string> Alert { get; set; }
        public DbHelper Db { get; set; }
        //public SaveTableToDbCommand()
        //{
        //}
        public IDisplayProgress progressDisplayHelper { get; set; }
        public DataSet TargetDataset { get; internal set; }

        public IEnumerable<string> Execute()
        {
            //we copy to server
            var table = TargetDataset.Tables[0];
            var targetTable = table.TableName;
            //we create the table
            var builder = new StringBuilder();
            foreach (DataColumn dc in table.Columns)
            {
                builder.AppendFormat("");
            }

            var res =
                string.Format("create table {0} ({1})", targetTable,
                string.Join(",",
                (
            from DataColumn dc in table.Columns
            select dc.ColumnName + " varchar(250)")));

            //initialise db
            
            int recordsImported = -5;
            using (var transaction = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(2, 0, 0)))
            {
                //and create the temp table
                Db.ExecSql(res);

                //use bulkcopy to write table to db
                Db.WriteTableToDb(targetTable, table);

                //we check how many records in the table
                recordsImported = Db.GetScalar("select count(*) from " + targetTable);

                transaction.Complete();
            }
            //MessageBox.Show("Records will be imported " + recordsImported);
            return new List<string>();
        }
    }

}
