using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ZambiaDataManager.Storage
{
    public class McspLngDataMergeCommand : BaseMergeCommand
    {
        public string TargetView { get; internal set; }

        protected override void DoMerge()
        {
            mergeWebData();
        }

        void mergeWebData()
        {
            var dbHelper = Db;
            IsInError = false;
            var sql = string.Empty;

            //step 1. Convert all values into standard strongly referenced values. Save to table temptable_1
            var newTempTableName = DestinationTable;
            sql = @"if(object_id('{0}') is not null) drop table {0};";
            dbHelper.ExecSql(string.Format(sql, newTempTableName));
            sql = @"select * into {0} from {1}";
            dbHelper.ExecSql(string.Format(sql, newTempTableName, TempTableName));

            sql = @"alter view view_lng_{0} as select * from {1}";
            dbHelper.ExecSql(string.Format(sql, TargetView, newTempTableName));
            
            //we clean up
            sql = @"if object_id('{0}') is not null drop table {0};";
            dbHelper.ExecSql(string.Format(sql, TempTableName));
        }
    }
}
