using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace db_resyncer.Model
{
    public static class TableHelper
    {
        internal static List<DbTable> getTableProcessors()
        {
            var tableNames = new List<string>() {
"AgeGroupLookUp",
"AgeGroupLookupAlternate",
"DistrictLookUp",
"FacilityData",
"FacilityLookUp",
"FacilityServiceLookup",
"FacilityTypeLookup",
"GenderLookUp",
"IndicatorLookup",
"MonthLookUp",
"ProgramAreaLookUp",
"ProvinceLookUp",
"YearLookUp" };

            var dbFields = new List<dbField>() {new dbField(){table_name="accessFacilityDataTemp",  column_name="FacilityIndex",    position=1, data_type="int"},
new dbField(){table_name="accessFacilityDataTemp",  column_name="Indicator",    position=2, data_type="string"},
new dbField(){table_name="accessFacilityDataTemp",  column_name="ReferenceYear",    position=3, data_type="int"},
new dbField(){table_name="accessFacilityDataTemp",  column_name="ReferenceMonth",   position=4, data_type="int"},
new dbField(){table_name="accessFacilityDataTemp",  column_name="Sex",  position=5, data_type="int"},
new dbField(){table_name="accessFacilityDataTemp",  column_name="AgeGroup", position=6, data_type="int"},
new dbField(){table_name="accessFacilityDataTemp",  column_name="Number",   position=7, data_type="int"},
new dbField(){table_name="accessFacilityDataTemp",  column_name="FacilityName", position=8, data_type="string"},
new dbField(){table_name="AgeGroupLookUp",  column_name="AgeGroupID",   position=1, data_type="int"},
new dbField(){table_name="AgeGroupLookUp",  column_name="AgeGroupName", position=2, data_type="string"},
new dbField(){table_name="AgeGroupLookupAlternate", column_name="Id",   position=1, data_type="int"},
new dbField(){table_name="AgeGroupLookupAlternate", column_name="AgeGroup", position=2, data_type="string"},
new dbField(){table_name="AgeGroupLookupAlternate", column_name="AgeGroupId",   position=3, data_type="int"},
new dbField(){table_name="DistrictLookUp",  column_name="DistrictID",   position=1, data_type="int"},
new dbField(){table_name="DistrictLookUp",  column_name="DistrictName", position=2, data_type="string"},
new dbField(){table_name="DistrictLookUp",  column_name="ProvinceID",   position=3, data_type="int"},
new dbField(){table_name="FacilityData",    column_name="Id",   position=1, data_type="int"},
new dbField(){table_name="FacilityData",    column_name="FacilityIndex",    position=2, data_type="int"},
new dbField(){table_name="FacilityData",    column_name="IndicatorSerial",  position=3, data_type="int"},
new dbField(){table_name="FacilityData",    column_name="ReferenceYear",    position=4, data_type="int"},
new dbField(){table_name="FacilityData",    column_name="ReferenceMonth",   position=5, data_type="int"},
new dbField(){table_name="FacilityData",    column_name="GenderId", position=6, data_type="int"},
new dbField(){table_name="FacilityData",    column_name="AgeGroupId",   position=7, data_type="int"},
new dbField(){table_name="FacilityData",    column_name="IndicatorValue",   position=8, data_type="float"},
new dbField(){table_name="facilityDataTemp",    column_name="IndicatorCode",    position=1, data_type="string"},
new dbField(){table_name="facilityDataTemp",    column_name="Attribute",    position=2, data_type="string"},
new dbField(){table_name="facilityDataTemp",    column_name="Value",    position=3, data_type="float"},
new dbField(){table_name="facilityDataTemp",    column_name="HmisCode", position=4, data_type="string"},
new dbField(){table_name="facilityDataTemp",    column_name="ReferenceYear",    position=5, data_type="int"},
new dbField(){table_name="facilityDataTemp",    column_name="ReferenceMonth",   position=6, data_type="int"},
new dbField(){table_name="FacilityLookUp",  column_name="FacilityIndex",    position=1, data_type="int"},
new dbField(){table_name="FacilityLookUp",  column_name="FacilityID",   position=2, data_type="string"},
new dbField(){table_name="FacilityLookUp",  column_name="FacilityName", position=3, data_type="string"},
new dbField(){table_name="FacilityLookUp",  column_name="DistrictID",   position=4, data_type="int"},
new dbField(){table_name="FacilityLookUp",  column_name="FacilityTypeID",   position=5, data_type="int"},
new dbField(){table_name="FacilityLookUp",  column_name="FacilityName_JHPEIGO", position=6, data_type="string"},
new dbField(){table_name="FacilityServiceLookup",   column_name="FacilityServiceID",    position=1, data_type="int"},
new dbField(){table_name="FacilityServiceLookup",   column_name="FacilityService",  position=2, data_type="string"},
new dbField(){table_name="FacilityTypeLookup",  column_name="FacilityTypeID",   position=1, data_type="int"},
new dbField(){table_name="FacilityTypeLookup",  column_name="FacilityType", position=2, data_type="string"},
new dbField(){table_name="GenderLookUp",    column_name="GenderID", position=1, data_type="int"},
new dbField(){table_name="GenderLookUp",    column_name="Gender",   position=2, data_type="string"},
new dbField(){table_name="GenderLookUp",    column_name="GenderLongName",   position=3, data_type="string"},
new dbField(){table_name="IndicatorLookup", column_name="IndicatorSerial",  position=1, data_type="int"},
new dbField(){table_name="IndicatorLookup", column_name="IndicatorID",  position=2, data_type="string"},
new dbField(){table_name="IndicatorLookup", column_name="IndicatorDescription", position=3, data_type="string"},
new dbField(){table_name="IndicatorLookup", column_name="zPosition",    position=4, data_type="int"},
new dbField(){table_name="IndicatorLookup", column_name="ProgramAreaID",    position=5, data_type="int"},
new dbField(){table_name="MonthLookUp", column_name="MonthID",  position=1, data_type="int"},
new dbField(){table_name="MonthLookUp", column_name="MonthName",    position=2, data_type="string"},
new dbField(){table_name="MonthLookUp", column_name="Quarter",  position=3, data_type="int"},
new dbField(){table_name="ProgramAreaLookUp",   column_name="ProgramAreaID",    position=1, data_type="int"},
new dbField(){table_name="ProgramAreaLookUp",   column_name="ProgramArea",  position=2, data_type="string"},
new dbField(){table_name="ProgramAreaLookUp",   column_name="AlternameName",    position=3, data_type="string"},
new dbField(){table_name="ProvinceLookUp",  column_name="ProvinceID",   position=1, data_type="int"},
new dbField(){table_name="ProvinceLookUp",  column_name="ProvinceName", position=2, data_type="string"},
new dbField(){table_name="Switchboard Items",   column_name="SwitchboardID",    position=1, data_type="int"},
new dbField(){table_name="Switchboard Items",   column_name="ItemNumber",   position=2, data_type="int"},
new dbField(){table_name="Switchboard Items",   column_name="ItemText", position=3, data_type="string"},
new dbField(){table_name="Switchboard Items",   column_name="Command",  position=4, data_type="int"},
new dbField(){table_name="Switchboard Items",   column_name="Argument", position=5, data_type="string"},
new dbField(){table_name="tblFacilityDataEntryTemplate",    column_name="ID",   position=1, data_type="int"},
new dbField(){table_name="tblFacilityDataEntryTemplate",    column_name="lt29", position=10,    data_type="int"},
new dbField(){table_name="tblFacilityDataEntryTemplate",    column_name="lt49", position=11,    data_type="int"},
new dbField(){table_name="tblFacilityDataEntryTemplate",    column_name="gte50",    position=12,    data_type="int"},
new dbField(){table_name="tblFacilityDataEntryTemplate",    column_name="rowtotal", position=13,    data_type="int"},
new dbField(){table_name="tblFacilityDataEntryTemplate",    column_name="IndicatorCode",    position=2, data_type="string"},
new dbField(){table_name="tblFacilityDataEntryTemplate",    column_name="Indicator",    position=3, data_type="string"},
new dbField(){table_name="tblFacilityDataEntryTemplate",    column_name="lt1",  position=4, data_type="int"},
new dbField(){table_name="tblFacilityDataEntryTemplate",    column_name="lt4",  position=5, data_type="int"},
new dbField(){table_name="tblFacilityDataEntryTemplate",    column_name="lt9",  position=6, data_type="int"},
new dbField(){table_name="tblFacilityDataEntryTemplate",    column_name="lt14", position=7, data_type="int"},
new dbField(){table_name="tblFacilityDataEntryTemplate",    column_name="lt19", position=8, data_type="int"},
new dbField(){table_name="tblFacilityDataEntryTemplate",    column_name="lt24", position=9, data_type="int"},
new dbField(){table_name="tblzNull",    column_name="ItemID",   position=1, data_type="int"},
new dbField(){table_name="tblzNull",    column_name="ItemName", position=2, data_type="string"},
new dbField(){table_name="tblzNull",    column_name="ItemName2",    position=3, data_type="string"},
new dbField(){table_name="YearLookUp",  column_name="YearID",   position=1, data_type="int"},
new dbField(){table_name="YearLookUp",  column_name="YearName", position=2, data_type="int"},
new dbField(){table_name="YearLookUp",  column_name="YPosition",    position=3, data_type="int"},
            };

            var tables = (from table in tableNames select new DbTable() { Name = table, FieldList = dbFields.Where(t => t.table_name == table).ToList() }).ToList();
            return tables;
        }
    }

    public class DbTable
    {
        public string Name { get; set; }
        public List<dbField> FieldList { get; set; }
    }

    public class dbField
    {
        public string table_name { get; set; }
        public string column_name { get; set; }
        public int position { get; set; }
        public string data_type { get; set; }
    }
}
