using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace HPVExcelReader
{
    public class GetExcelAsDataTable2 : ExcelWorksheetReaderBase, IQueryHelper<Dictionary<string, List<string>>>
    {
        public bool addHeaderRow { get; set; }
        public Microsoft.Office.Interop.Excel.Application excelApp = null;
        public bool IsInError { get; set; }
        //public DbHelper Db { get; set; }
        public List<string> expectedWorksheetNames = new List<string>() {
            "2a. Primary Schools","2b. Secondary Schools",
            "3. Communities ","4. Vaccine and supplies and M&E"};
        public List<string> firstHeader = new List<string>() {
            "Name of Primary  School",
            "Name of Secondary  School",
            "Name of Village or Township","Name of Health facility" };
        public List<string> preferredWorksheetname = new List<string>() {
            "Primary  School",
            "Secondary  School",
            "Community Name","Health facility" };
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

        public static string getCellValue(Range xlrange, int rowId, int colmnId, string valueIfNull = "")
        {
            var cellvalue = Convert.ToString(xlrange[rowId, colmnId].Value);
            return cellvalue == null ? (valueIfNull == "" ? string.Empty : valueIfNull) : cellvalue.ToString().Trim();
        }
        public class headerIndexes
        {
            public int worksheetId { get; set; }
            public string worksheetName { get; set; }
            public int firstCellRowId { get; set; }
            public int firstCellColumnId { get; set; }
            public Range range { get; set; }
        }

        public class dataObject
        {
            public string filename { get; set; }
            public string worksheet { get; set; }
            public List<string> values { get; set; }
            //public int firstCellRowId { get; set; }
            //public int firstCellColumnId { get; set; }
            //public Range range { get; set; }
        }

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
                if (!ds.ContainsKey(rangename))
                {
                    ds[rangename] = new List<string>();
                }
                var datavalues = ds[rangename];

                var rowid = rng.firstCellRowId;
                var rowmax = rng.range.Rows.Count;

                var columnid = rng.firstCellColumnId;
                var colmax = maxColumnsExpected[rng.worksheetId] + columnid;

                //first row is for the headers                
                for(var rowindx = rowid; rowindx < rowmax; rowindx++)
                {
                    //we skip the first row
                    if (rowindx == rowid)
                    {
                        if (!addHeaderRow)
                            continue;
                    }

                    var builder = new StringBuilder();
                    builder.Append(Path.GetFileNameWithoutExtension(fileName));
                    string schoolname = string.Empty;
                    var quitrow = false;
                    for(var colmnindx= columnid; colmnindx < colmax; colmnindx++)
                    {
                        //we check if the column is blank, if so, we break inner loop
                        var cellValue = getCellValue(rng.range, rowindx, colmnindx, "9999").Replace(",", "-").Replace("\"", "-");
                        schoolname = cellValue;
                        var specialCase = "4. Vaccine and supplies and M&E";
                        //rowindx = rowid
                        if (colmnindx == columnid && (specialCase != rangename || (specialCase == rangename && rowindx != rowid + 1)))
                        {
                            if ("9999" == cellValue)
                            {
                                //we quit here
                                quitrow = true;
                                break;
                            }
                            //schoolname = cellValue.Replace(",", "-");
                            //builder.Append(cellValue);
                        }
                        builder.AppendFormat("\t{0}", cellValue);
                    }
                    if (quitrow)
                    {
                        //no need to add
                        break;
                    }
                    var str = builder.ToString();
                    if (!str.Contains("9999,9999,9999"))
                        datavalues.Add(builder.ToString());
                }
                
            }


            return ds;
        }
    }

}
