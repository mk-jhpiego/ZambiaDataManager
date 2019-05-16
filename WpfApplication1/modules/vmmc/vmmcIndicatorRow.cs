using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZambiaDataManager.modules.vmmc
{
    public class vmmcIndicatorRow : ISerialisableWebDataset
    {
        //we get the data for each indicator
        public string provinceName;
        public string districtName;
        public string facilityName;
        public string reportMonth;
        public string reportYear;
        public string cleanname;
        public string indicator;
        public int age10to14;
        public int age15to19;
        public int age20to24;
        public int age25to29;
        public int age30to49;
        public int age50plus;

        public int age30to34;
        public int age35to39;
        public int age40to49;
        public int age40to44;
        public int age45to49;

        public string user;

        public DataTable getTable()
        {
            var table = new DataTable();
            var fields = new List<string>{"provinceName",
            "districtName",
            "facilityName",
            "reportYear",
            "reportMonth",
            //"cleanname",
            "indicator",
            "age10to14",
            "age15to19",
            "age20to24",
            "age25to29",
            "age30to34",
            "age35to39",
            "age40to44",
            "age45to49",
            "age40to49",
            "age50plus",
            "user_name","age30to49"};
            fields.ForEach(t => table.Columns.Add(t));
            return table;
        }

        public DataRow toRow(DataRow row)
        {
            row["provinceName"] = provinceName;
            row["districtName"] = districtName;
            row["facilityName"] = facilityName;
            row["reportYear"] = reportYear;
            row["reportMonth"] = reportMonth;
            //row["cleanname"] = cleanname;
            row["indicator"] = indicator;
            row["age10to14"] = age10to14;
            row["age15to19"] = age15to19;
            row["age20to24"] = age20to24;
            row["age25to29"] = age25to29;
            row["age30to49"] = age30to49;

            row["age30to34"] = age30to34;
            row["age35to39"] = age35to39;

            row["age40to44"] = age40to44;
            row["age45to49"] = age45to49;
            row["age40to49"] = age40to49;

            row["age50plus"] = age50plus;
            row["user_name"] = user;
            return row;
        }
    }
}
