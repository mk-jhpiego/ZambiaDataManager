using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace HPVExcelReader
{
    public class GetMicroplans : ExcelWorksheetReaderBase, IQueryHelper<Dictionary<string, List<string>>>
    {
        public List<string> expectedWorksheetNames = new List<string>() {
            "2a. Primary Schools","2b. Secondary Schools",
            "3. Communities "};
        public List<string> firstHeader = new List<string>() {
            "Name of Primary  School",
            "Name of Secondary  School",
            "Name of Village or Township"};
        public List<string> preferredWorksheetname = new List<string>() {
            "Primary  School",
            "Secondary  School",
            "Community Name"};
        public List<int> maxColumnsExpected = new List<int>() { 2, 2, 3, 3 };

        public Dictionary<string, List<string>> Execute()
        {

            //ProjectName projectName
            Dictionary<string, List<string>> toReturn = null;
            try
            {                
                var res = ImportData();
                //add location details here
                if (IsInError)
                    return null;
                toReturn = res;
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("{0} --ERROR-- {1}", fileName, ex.Message));
                //MessageBox.Show(ex.Message);
            }
            finally
            {
                if (excelApp != null)
                {
                    excelApp.Quit();
                }
            }
            return toReturn;
        }

        protected static RowColmnPair GetRequiredValueCell(string firstHeader, Range xlrange)
        {
            int colCount = xlrange.Columns.Count;
            colCount = colCount > 15 ? 12 : colCount;
            int row = -1, colmn = -1, colmn2 = -1;
            var matchfound = false;
            var maxDepthSearchRows = 10;
            //var cleanAgeDisaggs = dataElement.getCleanAgeDisaggregations();

            for (var rowId = 1; rowId <= maxDepthSearchRows; rowId++)
            {
                for (var colmnId = 1; colmnId <= colCount; colmnId++)
                {
                    var rawAgeValue = getCellValue(xlrange, rowId, colmnId);
                    if (string.IsNullOrWhiteSpace(rawAgeValue) || rawAgeValue.Length > 40) continue;

                    var cleanAgeValue = rawAgeValue;
                    //no longer searching if its thefirst item, since we are going from Column 1
                    if (firstHeader == cleanAgeValue)
                    {
                        //we've found our column, time to find where the data begibs, lets find the corresponding indicator
                        //we'll scan for columns from rowid to perhaps 5 places, and starting from column 0
                        row = rowId;
                        colmn = colmnId;
                        matchfound = true;
                        break;
                    }
                }
                if (matchfound) break;
            }
            return new RowColmnPair(row, colmn, colmn2);
        }

        public enum SchoolType
        {
            Primary = 0, Secondary = 1, Community = 2
        }

        List<string> getDbfields(SchoolType sctype)
        {
            List<string> defaultFields = null;
            switch (sctype)
            {
                case SchoolType.Primary:
                    defaultFields = new List<string>() { "id", "src_filepath", "src_file", "School Name", "Girls Aged 14" };
                    break;
                case SchoolType.Secondary:
                    defaultFields = new List<string>() { "id", "src_filepath", "src_file", "School Name", "Girls Aged 14" };
                    break;
                case SchoolType.Community:
                    defaultFields = new List<string>() { "id", "src_filepath", "src_file", "Village or Township", "Health Facilility Name", "Girls out of School", "Name CHW", "Phone CHW" };
                    break;
            }
            //drow["src_filepath"] = fileName;
            //drow["src_file"] = Path.GetFileName(fileName);
            return (from field in defaultFields
                    let cname = field.ToLowerInvariant().Trim().Replace(" ", "_").Replace(".", "")
                    select cname).ToList();
        }
        string getTableName(SchoolType sctype)
        {
            var name = string.Empty;

            switch (sctype)
            {
                case SchoolType.Primary:
                    name = "girls_primary";
                    break;
                case SchoolType.Secondary:
                    name = "girls_secondary";
                    break;
                case SchoolType.Community:
                    name = "girls_community";
                    break;
            }
            return name;
        }

        public const int MAX_BLANKROWS = 5;
        private Dictionary<string, List<string>> ImportData()
        {
            var ds = new Dictionary<string, List<string>>();
            //var dataObjects = new List<dataObject>();
            var fileShortName = Path.GetFileNameWithoutExtension(fileName);
            var workbook = excelApp.Workbooks.Open(fileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            //we get the other data
            var worksheetCount = workbook.Sheets.Count;
            var worksheetNames = new Dictionary<string, string>();
            
            for (var indx = 1; indx <= worksheetCount; indx++)
            {
                var worksheetName = ((Worksheet)(workbook.Sheets[indx])).Name;
                worksheetNames.Add(worksheetName.Trim().ToLowerInvariant(), worksheetName);
            }

            var sz = expectedWorksheetNames.Count();
            var workbookErrors = new StringBuilder();
            var hasErrors = false;
            var buildr = new StringBuilder();
            var availableRanges = new List<headerIndexes>();

            for (var ix = 0; ix < sz; ix++)
            {
                Console.WriteLine(ix);
                var name = expectedWorksheetNames[ix];
                var firstheader = firstHeader[ix];

                var cleanname = name.Trim().ToLowerInvariant();
                var hasName = worksheetNames.ContainsKey(cleanname);
                if (!hasName)
                {
                    if (!hasErrors)
                    {
                        workbookErrors.AppendLine(fileName);
                    }
                    hasErrors = true;
                    workbookErrors.AppendLine("\t"+name);
                    //File.AppendAllText("failing.txt", fileName + "\n");
                    Console.WriteLine("Missing worksheet,\t{0},\t{1}}", name, fileName);
                }
                else
                {
                    var xlrange = ((Worksheet)workbook.Sheets[worksheetNames[cleanname]]).UsedRange;
                    var rowCount = xlrange.Rows.Count;
                    var colCount = xlrange.Columns.Count;
                    var matched = GetRequiredValueCell(firstheader, xlrange);
                    availableRanges.Add(new headerIndexes
                    {
                        worksheetId = ix,
                        firstCellColumnId = matched.Column,
                        firstCellRowId = matched.Row,
                        worksheetName = worksheetNames[cleanname],
                        range= xlrange
                    });
                    //Console.WriteLine(string.Format("{2},\t{0},\t{1}", matched.Row, matched.Column, name));
                    buildr.AppendLine(string.Format("{0}\t{1}\t{2}\t{3}\t{4}", name, matched.Row, matched.Column, Path.GetFileNameWithoutExtension(fileName), fileName));

                }
            }
            if (hasErrors)
            {
                File.AppendAllText("failing.txt", workbookErrors.ToString() + "\n");
                Console.WriteLine(workbookErrors.ToString());
                return null;
            }
            
            File.AppendAllText("headermatches.csv", buildr.ToString());

            //var breakAt = 1;
            
            foreach (var rng in availableRanges)
            {
                var rangename = expectedWorksheetNames[rng.worksheetId];
                var sctype = (SchoolType)rng.worksheetId;

                if (!ds.ContainsKey(rangename))
                {
                    ds[rangename] = new List<string>();
                }
                var datavalues = ds[rangename];

                var fields = getDbfields(sctype);
                var brd = new StringBuilder();
                brd.AppendFormat("create table [{0}](", getTableName(sctype));
                fields.ForEach(x => brd.AppendFormat("[{0}] varchar(255),", x));
                var asString = brd.ToString();
                var dds = new System.Data.DataSet();
                var dt = new System.Data.DataTable(getTableName(sctype));
                dds.Tables.Add(dt);
                foreach (var field in fields)
                {
                    dt.Columns.Add(field);
                }

                var rowid = rng.firstCellRowId;
                var rowmax = rng.range.Rows.Count;

                var columnid = rng.firstCellColumnId;
                var colmax = maxColumnsExpected[rng.worksheetId] + columnid;

                //first row is for the headers      
                var blankRowCount = 0;
                for (var rowindx = rowid; rowindx < rowmax; rowindx++)
                {
                    //we skip the first row
                    if (rowindx == rowid)
                    {
                        if (!addHeaderRow)
                            continue;
                    }

                    var drow = dt.NewRow();
                    drow["id"] = rowindx;
                    drow["src_filepath"] = fileName;
                    drow["src_file"] = Path.GetFileName(fileName);

                    var builder = new StringBuilder();
                    builder.Append(Path.GetFileNameWithoutExtension(fileName));
                    string schoolname = string.Empty;
                    var quitrow = false;
                    for(var colmnindx= columnid; colmnindx < colmax; colmnindx++)
                    {
                        if (colmnindx - columnid + 3 >= fields.Count)
                        {
                            break;
                        }
                        //we check if the column is blank, if so, we break inner loop
                        var cellValue = getCellValue(rng.range, rowindx, colmnindx, "9999").checkLength2();
                        //var cellValue = getCellValue(rng.range, rowindx, colmnindx, "9999").Replace(",", "-").Replace("\"", "-");
                        schoolname = cellValue;
                        var specialCase = "4. Vaccine and supplies and M&E----";
                        //rowindx = rowid
                        if (colmnindx == columnid && (specialCase != rangename || (specialCase == rangename && rowindx != rowid + 1)))
                        {
                            if ("9999" == cellValue)
                            {
                                blankRowCount++;
                                if (blankRowCount >= MAX_BLANKROWS)
                                {
                                    //we quit here
                                    quitrow = true;
                                }
                                break;
                            }
                            else
                            {
                                blankRowCount = 0;
                            }
                            //if ("9999" == cellValue)
                            //{
                            //    //we quit here
                            //    quitrow = true;
                            //    break;
                            //}

                        }
                        var colmnname = fields[colmnindx - columnid + 3];
                        drow[colmnname] = (cellValue == "9999" || cellValue == "-2146826273" || cellValue == "-2146826265") ? "0" : cellValue;
                        builder.AppendFormat("\t{0}", cellValue);
                    }
                    if (quitrow)
                    {
                        //no need to add
                        break;
                    }
                    var str = builder.ToString();
                    //datavalues.Add(builder.ToString());                    
                    if (!str.Contains("9999,9999,9999"))
                    {
                        dt.Rows.Add(drow);
                        datavalues.Add(builder.ToString());
                    }
                }
                //we save
                dt.AcceptChanges();
                var columnList = new StringBuilder();
                dt.Columns.Cast<DataColumn>().ToList().ForEach(c => columnList.AppendFormat("{0},", c));

                db.WriteTableToDb(dt, getTableName(sctype));
            }
            return ds;
        }
    }
}
