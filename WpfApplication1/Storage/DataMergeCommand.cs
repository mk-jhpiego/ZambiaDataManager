using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ZambiaDataManager.Storage
{
    public class DataMergeCommand : BaseMergeCommand
    {
        protected override void DoMerge()
        {
            var dbHelper = Db;
            IsInError = false;
            var sql = string.Empty;
            //steps
            //read excel file to table facilityDataTemp
            ///x. copyDataFromExcel "facilityDataTemp", strSourceDatabase2

            //we check if all age categories are mapped to existing ids
            sql = "select distinct AgeGroup from {0} except select AgeGroup from AgeGroupLookupAlternate";
            var values = dbHelper.GetListText(string.Format(sql, TempTableName));

            if (values.Count > 0)
            {
                //show error and return
                IsInError = true;
                MessageBox.Show("The following categories could not be processed: " + string.Join(",", values));
                return;
            }

            //step 1. Convert all values into standard strongly referenced values. Save to table temptable_1
            var newTempTableName = TempTableName + "_1";
            sql = @"
select f.FacilityIndex, 
d.FacilityName FacilityId,
il.IndicatorSerial,
y.YearID,
m.MonthID ReferenceMonth,
g.GenderId Sex,
a.AgeGroupId,
IndicatorValue Value
into {1}
 From {0} d
 join FacilityLookupAll f on d.FacilityName = rtrim(ltrim(f.FacilityHmisCode))
 join YearLookUp y on d.ReportYear = y.YearName
 join MonthLookUp m on d.ReportMonth = m.MonthName
 join AgeGroupLookupAlternate a on d.AgeGroup = a.AgeGroup
 join GenderLookUp g on d.Sex = g.GenderLongName
 join [dbo].[IndicatorLookup] il on d.IndicatorId = il.IndicatorId
";
            dbHelper.ExecSql(string.Format(sql, TempTableName, newTempTableName));

            //we check if the data we have is unique or already exists
            sql = @"
select count(*) as tCount from {0} f 
join FacilityData t on 
f.facilityindex = t.FacilityIndex and f.YearId = t.ReferenceYear
 and f.ReferenceMonth = t.ReferenceMonth
";

            MessageBoxResult dlgRes;
            var recCount = dbHelper.GetScalar(string.Format(sql, newTempTableName));
            if (recCount > 0)
            {
                dlgRes = MessageBox.Show("Records for this facility already exist on the server. Do you want to stop this process?", "Confirm Action", MessageBoxButton.YesNo);
                if (dlgRes == MessageBoxResult.Yes)
                {
                    //we abort
                    return;
                }
                else
                {
                    dlgRes = MessageBox.Show("Matching records on the server will be deleted. Do you want to delete the records on the server and replace them with this?", "Confirm Further Action", MessageBoxButton.YesNo);
                    if (dlgRes != MessageBoxResult.Yes)
                    {
                        //we abort
                        return;
                    }
                }

                //we delete all matching records
                sql = @"
 delete from FacilityData where Id in (
 select distinct Id from FacilityData f 
 join
 (select distinct facilityindex, YearId, ReferenceMonth from  {0}) t 
 on f.facilityindex = t.FacilityIndex and f.ReferenceYear = t.YearId
 and f.ReferenceMonth = t.ReferenceMonth)";
                dbHelper.ExecSql(string.Format(sql, newTempTableName));
            }

            //we import the data
            sql = @"
INSERT INTO [dbo].[FacilityData]
([FacilityIndex],[IndicatorSerial],[ReferenceYear]
,[ReferenceMonth],[GenderId],[AgeGroupId],[IndicatorValue])
SELECT 
[FacilityIndex],[IndicatorSerial],[YearID],
[ReferenceMonth],[Sex] as GenderId,[AgeGroupId],
[Value] as IndicatorValue FROM {0}";
            dbHelper.ExecSql(string.Format(sql, newTempTableName));

            //we clean up
            sql = "drop table {0};drop table {1}; ";
            dbHelper.ExecSql(string.Format(sql, newTempTableName, TempTableName));
        }
    }

    public class ProjectFinanceMergeCommand : BaseMergeCommand
    {
        protected override void DoMerge()
        {
            var dbHelper = Db;
            IsInError = false;
            var sql = string.Empty;
            //steps
            //read excel file to table facilityDataTemp
            ///x. copyDataFromExcel "facilityDataTemp", strSourceDatabase2

            //we check if all age categories are mapped to existing ids
            sql = "select distinct ProjectMatchKey from {0} except select ProjectCode from ProjectCodes";
            var values = dbHelper.GetListText(string.Format(sql, TempTableName));
            if (values.Count > 0)
            {
                //show error and return
                IsInError = true;
                MessageBox.Show("The following IONs could not be processed: " + string.Join(",", values));
                return;
            }

            sql = "select distinct IndicatorId from {0} except select GlCode from GlCodes";
            values = dbHelper.GetListText(string.Format(sql, TempTableName));
            if (values.Count > 0)
            {
                //show error and return
                IsInError = true;
                MessageBox.Show("The following GL Codes could not be processed: " + string.Join(",", values));
                return;
            }

            //step 1. Convert all values into standard strongly referenced values. Save to table temptable_1
            var newTempTableName = TempTableName + "_1";
            //insert into ProjectSpendingDetails(ProjectId, ReferenceYear, ReferenceMonth, GlCodeId, TotalSpending, OfficeAllocation, DirectCost)
            sql = @"
select ProjectId, YearId, MonthId, GlCodeId, TotalCost TotalSpending, OfficeAllocation, DirectCost
into {1}
from {0} main
join GlCodes gl on IndicatorId = gl.GlCode
join ProjectCodes pcode on ProjectMatchKey = pcode.ProjectCode
join YearLookUp yr on main.ReportYear = yr.YearName
join MonthLookUp mnth on main.ReportMonth = mnth.[MonthName]
";
            dbHelper.ExecSql(string.Format(sql, TempTableName, newTempTableName));

            //we check if the data we have is unique or already exists
            sql = " select count(*) as tCount from {0} f join ProjectSpendingDetails t on f.YearId = t.YearId and f.MonthId = t.MonthId";

            MessageBoxResult dlgRes;
            var recCount = dbHelper.GetScalar(string.Format(sql, newTempTableName));
            var deleteMatchingRecords = false;
            CommandParam deleteDataParams = null;
            if (recCount > 0)
            {
                dlgRes = MessageBox.Show("Records for the month specified already exist on the server. Do you want to stop this process?", "Confirm Action", MessageBoxButton.YesNo);
                if (dlgRes == MessageBoxResult.Yes)
                {
                    //we abort
                    return;
                }
                else
                {
                    dlgRes = MessageBox.Show("Matching records on the server will be deleted. Do you want to delete the records on the server and replace them with this?", "Confirm Further Action", MessageBoxButton.YesNo);
                    if (dlgRes != MessageBoxResult.Yes)
                    {
                        //we abort
                        return;
                    }
                }

                //we delete all matching records
                deleteMatchingRecords = true;
                sql = " select distinct YearId, MonthId from {0}";
                var rowdata = dbHelper.GetTable(string.Format(sql, newTempTableName)).Rows[0];
                deleteDataParams = new CommandParam().Add("@yearId", rowdata[0]).Add("@monthId", rowdata[1]);
            }

            //we import the data
            if (deleteMatchingRecords)
            {
                //deleteDataParams
                sql = @"
begin transaction;
delete from ProjectSpendingDetails where YearId = @yearId and MonthId = @monthId;
 insert into ProjectSpendingDetails(ProjectId, YearId, MonthId, GlCodeId, TotalSpending, OfficeAllocation, DirectCost)
select ProjectId, YearId, MonthId, GlCodeId, TotalSpending, OfficeAllocation, DirectCost
from {0};
commit transaction;";
                dbHelper.ExecSql(string.Format(sql, newTempTableName), deleteDataParams);
            }
            else
            {
                sql = @"
insert into ProjectSpendingDetails(ProjectId, YearId, MonthId, GlCodeId, TotalSpending, OfficeAllocation, DirectCost)
select ProjectId, YearId, MonthId, GlCodeId, TotalSpending, OfficeAllocation, DirectCost
from {0};";
                dbHelper.ExecSql(string.Format(sql, newTempTableName));
            }

            //we clean up
            sql = "drop table {0};drop table {1}; ";
            dbHelper.ExecSql(string.Format(sql, newTempTableName, TempTableName));
        }

    }

    public class BaseMergeCommand : IQueryHelper<IEnumerable<string>>
    {
        public Action<string> Alert { get; set; }
        public DbHelper Db { get; set; }

        public bool IsInError { get; set; }
        public string DestinationTable { get; set; }
        public string TempTableName { get; set; }

        public IDisplayProgress progressDisplayHelper { get; set; }
        public IEnumerable<string> Execute()
        {
            //
            DoMerge();
            return new List<string>();
        }

        protected virtual void DoMerge()
        {
        }
    }
}
