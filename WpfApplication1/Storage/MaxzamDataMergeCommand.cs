using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ZambiaDataManager.Storage
{
    public class MaxzamDataMergeCommand : BaseMergeCommand
    {
        internal Dictionary<string, string> destinationTempNames { get; set; }

        protected override void DoMerge()
        {
            var dbHelper = Db;
            IsInError = false;
            var dropTableSql = string.Empty;
            var sql = string.Empty;
            foreach (var targettable in destinationTempNames.Keys)
            {
                var temptable = destinationTempNames[targettable];
                //we clear the current database table
                sql = "delete from {0}";
                dbHelper.ExecSql(string.Format(sql, targettable));

                //we get the list of expected columns for target table
                //select column_name from information_schema.columns where table_name = '${0}'
                sql = @"
DECLARE @listStr VARCHAR(MAX)
SELECT @listStr = COALESCE(@listStr+',' , '') + '['+column_name+']'
from information_schema.columns 
where table_name = '{0}'
SELECT @listStr";
                var colmns = dbHelper.GetListText(string.Format(sql, targettable)).FirstOrDefault();
                if (string.IsNullOrWhiteSpace(colmns))
                {
                    throw new Exception(string.Format("Could not get information_schema.columns for {0}", targettable));
                }

                //we copy from the temp
                sql = "insert into {0}({2}) select {2} from {1}";
                dbHelper.ExecSql(string.Format(sql, targettable, temptable, colmns));

                dropTableSql += string.Format("drop table {0};", temptable);
            }
            dbHelper.ExecSql(dropTableSql);
            
        }

    }    
}
