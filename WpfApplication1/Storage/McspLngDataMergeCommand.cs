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
            var newTempTableName = TempTableName + "_1";
            sql = @"
with t as (
select FacilityIndex, IndicatorSerial, indicator, 
YearId ReferenceYear,convert(int, reportMonth) ReferenceMonth,2 GenderId, 
age10to14, age15to19,age20to24, age25to29, age30to49, age50plus
From {0} d
join dbo.fn_getWebFacilities() wf on d.facilityname = wf.fcleanname
join FacilityList f on wf.Facility = f.FacilityName
join dbo.fn_get_mcnames() mci on indicator = mci.shortname
join YearLookUp yr on d.reportYear = yr.YearName
)
select * into {1} from (
select FacilityIndex, IndicatorSerial, ReferenceYear, ReferenceMonth,
GenderId, 4 AgeGroupId, age10to14 IndicatorValue
From t union 
select FacilityIndex, IndicatorSerial, ReferenceYear, ReferenceMonth,
GenderId, 5 AgeGroupId, age15to19 IndicatorValue
From t union 
select FacilityIndex, IndicatorSerial, ReferenceYear, ReferenceMonth,
GenderId, 6 AgeGroupId, age20to24 IndicatorValue
From t union 
select FacilityIndex, IndicatorSerial, ReferenceYear, ReferenceMonth,
GenderId, 7 AgeGroupId, age25to29 IndicatorValue
From t union 
select FacilityIndex, IndicatorSerial, ReferenceYear, ReferenceMonth,
GenderId, 8 AgeGroupId, age30to49 IndicatorValue
From t union 
select FacilityIndex, IndicatorSerial, ReferenceYear, ReferenceMonth,
GenderId, 9 AgeGroupId, age50plus IndicatorValue
From t)k;
";
            dbHelper.ExecSql(string.Format(sql, TempTableName, newTempTableName));

            //we check how many records came through
            sql = "select count(*) from {0}";
            var recCountOldTable = dbHelper.GetScalar(string.Format(sql, TempTableName));
            if (recCountOldTable == 0)
            {
                var results = MessageBox.Show(
                   "Press OK to close", "No  records to import", MessageBoxButton.OK);
                return;
            }

            var recCountNewTable = dbHelper.GetScalar(string.Format(sql, newTempTableName));
            if (recCountNewTable != (recCountOldTable*6))
            {
                var expcount = recCountOldTable * 6;
                var results = MessageBox.Show(
                   string.Format("Count of records about to be imported ({0}) does not match the expected number ({1}). Did you add some new sites? If so, add them to fn_getWebFacilities before proceeding. Do you want to continue?",
                   recCountNewTable, expcount), "Mismatch in records to be imported", MessageBoxButton.YesNo);
                if (results == MessageBoxResult.No)
                {
                    //we abort
                    return;
                }
            }

            //we check if the data we have is unique or already exists
            //first we add to a new table
            sql = @"
 with db as (select distinct FacilityIndex, ReferenceYear, ReferenceMonth from {0})
 select db.* into {1} From db
 join (
  select distinct FacilityIndex, ReferenceYear, ReferenceMonth from FacilityData
 )dn on  db.FacilityIndex = dn.FacilityIndex and 
 db.ReferenceYear = dn.ReferenceYear and db.ReferenceMonth = dn.ReferenceMonth
";
            var newTempTableName2 = TempTableName + "_2";
            dbHelper.ExecSql(string.Format(sql, newTempTableName, newTempTableName2));

            sql = "select count(*) recs from {0}";
            MessageBoxResult dlgRes;
            var recCount = dbHelper.GetScalar(string.Format(sql, newTempTableName2));
            if (recCount > 0)
            {
                dlgRes = MessageBox.Show("Some records already exist on the server. Do you want to stop this process?", "Confirm Action", MessageBoxButton.YesNo);
                if (dlgRes == MessageBoxResult.Yes)
                {
                    //we abort
                    return;
                }
                else
                {
                    dlgRes = MessageBox.Show("Matching records on the server will be deleted. Do you want to delete the records on the server and replace them with these?", "Confirm Further Action", MessageBoxButton.YesNo);
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
 join {0} t on f.facilityindex = t.FacilityIndex and f.ReferenceYear = t.ReferenceYear
 and f.ReferenceMonth = t.ReferenceMonth)";
                dbHelper.ExecSql(string.Format(sql, newTempTableName));
            }

            //we import the data
            sql = @"
INSERT INTO [dbo].[FacilityData]
([FacilityIndex],[IndicatorSerial],[ReferenceYear]
,[ReferenceMonth],[GenderId],[AgeGroupId],[IndicatorValue])
SELECT 
[FacilityIndex],[IndicatorSerial],[ReferenceYear],
[ReferenceMonth],2 as GenderId,[AgeGroupId],
IndicatorValue FROM {0}";
            dbHelper.ExecSql(string.Format(sql, newTempTableName));
            //recCount = dbHelper.GetScalar(string.Format(sql, newTempTableName));

            //we clean up
            sql = @"if object_id('{0}') is not null drop table {0};
if object_id('{1}') is not null drop table {1};
if object_id('{2}') is not null drop table {2};";
            dbHelper.ExecSql(string.Format(sql, newTempTableName, TempTableName, newTempTableName2));
        }
    }
}
