using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace HPVExcelReader
{
    //GetExcelAsDataTable4
    public class GetVaccinesHelper : ExcelWorksheetReaderBase, IQueryHelper<Dictionary<string, List<string>>>
    {        
        //public DbHelper Db { get; set; }
        public List<string> expectedWorksheetNames = new List<string>() {
            "4. Vaccine and supplies and M&E"};

        public List<string> firstHeader = new List<string>() { "No.", "Name of Health facility",
            "Number of schools in catchment area ", "Target for HPV (Total number of girls 14 Primary School + Secondary School + Out of School )",
            "Vaccines and supplies", "", "", "", "Total number of HPV Vaccine cards", "HPV Summary Sheets",
            "HPV Registers", "Type of Cold Chain required (Cold box or carrier)", "Cold Chain", "", "" };
        public List<string> preferredWorksheetname = new List<string>() {"Vaccines" };
        public const int MAX_BLANKROWS = 10;
        public List<int> maxColumnsExpected = new List<int>() { 15 };

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
            //colCount = colCount > 15 ? 12 : colCount;
            int row = -1, colmn = -1, colmn2 = -1;
            var matchfound = false;
            var maxDepthSearchRows = 50;
            //var cleanAgeDisaggs = dataElement.getCleanAgeDisaggregations();
            var blankSoFar = 1;
            for (var rowId = 1; rowId <= maxDepthSearchRows; rowId++)
            {
                for (var colmnId = 1; colmnId <= colCount; colmnId++)
                {
                    var rawAgeValue = getCellValue(xlrange, rowId, colmnId);
                    if (string.IsNullOrWhiteSpace(rawAgeValue) || rawAgeValue.Length > 65) continue;

                    blankSoFar = 0;
                    var cleanAgeValue = rawAgeValue.ToLowerInvariant();
                    //no longer searching if its thefirst item, since we are going from Column 1
                    if (firstHeader == cleanAgeValue || firstHeader == cleanAgeValue + ".")
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
            return new RowColmnPair(row, colmn, blankSoFar);
        }

        List<string> getDbfields()
        {
            //drow["src_filepath"] = fileName;
            //drow["src_file"] = Path.GetFileName(fileName);
            var fieldnames = new List<string>() { "id","src_filepath","src_file","numb", "Health facility",
                "Number of schools",
                "Target for HPV",
                "HPV vaccines doses", "HPV vaccines vials", "AD Syringes", "Safety boxes",
                "Vaccine cards", "Summary Sheets", "Registers",
                "CC Type Required", "CC Number available ",
                "CC Number Required", "CC Gap" };
            return (from field in fieldnames
                    let cname = field.ToLowerInvariant().Trim().Replace(" ", "_")
                    select cname).ToList();
        }

        private Dictionary<string, List<string>> ImportData()
        {
            var fields = getDbfields();
            var dds = new System.Data.DataSet();
            var dt = new System.Data.DataTable("vaccines_list");
            dds.Tables.Add(dt);
            foreach(var field in fields)
            {
                dt.Columns.Add(field);
            }

            var expectedSheetRows = new List<string>();
            expectedWorksheetNames.ForEach(t => expectedSheetRows.Add(t.Trim().ToLowerInvariant()));

            var ds = new Dictionary<string, List<string>>();
            //var dataObjects = new List<dataObject>();
            var fileShortName = Path.GetFileNameWithoutExtension(fileName);
            var workbook = excelApp.Workbooks.Open(fileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            //we get the other data
            var worksheetCount = workbook.Sheets.Count;
            var worksheetNames = new Dictionary<string, string>();
            var hasExpectedSheets = false;
            for (var indx = 1; indx <= worksheetCount; indx++)
            {
                var worksheetName = ((Worksheet)(workbook.Sheets[indx])).Name;
                var cleanName = worksheetName.Trim().ToLowerInvariant();
                if (expectedSheetRows.Contains(cleanName))
                {
                    hasExpectedSheets = true;
                    worksheetNames.Add(cleanName, worksheetName);
                    break;
                }                
            }

            if (!hasExpectedSheets)
            {
                return null;
            }
            //we open each worksheet and search for the required columns
            //var sz = expectedWorksheetNames.Count();
            var workbookErrors = new StringBuilder();

            var buildr = new StringBuilder();
            var availableRanges = new List<headerIndexes>();
            var ctr = 0;
            var skipSheets = new List<string>() { "inside cover pg", "cover", "instructions" };

            var includeFilenameInNoMatch = true;
            foreach (var wknames in worksheetNames)//(var ix = 0; ix < sz; ix++)
            {
                Console.WriteLine(wknames.Key);
                if (wknames.Key.Contains("cover") || wknames.Key.Contains(" map ") || wknames.Key.Contains("instructions") || skipSheets.Contains(wknames.Key))//skipSheets.Contains(wknames.Key) || 
                {
                    continue;
                }
                //var name = expectedWorksheetNames[ix];

                var firstheader = firstHeader[0].ToLowerInvariant();
                var xlrange = ((Worksheet)workbook.Sheets[wknames.Value]).UsedRange;
                var rowCount = xlrange.Rows.Count;
                var colCount = xlrange.Columns.Count;
                var matched = GetRequiredValueCell(firstheader, xlrange);
                if (matched.Row == -1 || matched.Column == -1)
                {
                    headersMatched = false;
                    if (includeFilenameInNoMatch)
                    {
                        includeFilenameInNoMatch = false;
                        File.AppendAllText(@"ds\vaccines\nomatch.txt", string.Format("{2}{1}{0}{1}{2}{1}", fileName, Environment.NewLine, "**********"));
                    }
                    File.AppendAllText(@"ds\vaccines\nomatch.txt", string.Format("{2}\t{0}{1}",
                        wknames.Value, Environment.NewLine, matched.Column2 == 1 ? "BLANK" : ""));
                    Console.WriteLine(string.Format("Unmatched sheet {0}", wknames.Value));
                }
                else
                {
                    availableRanges.Add(new headerIndexes
                    {
                        worksheetId = ctr++,
                        firstCellColumnId = matched.Column,
                        firstCellRowId = matched.Row,
                        worksheetName = wknames.Value,
                        rangeName = wknames.Value.ToLowerInvariant().Trim(),
                        range = xlrange
                    });
                    buildr.AppendLine(string.Format("{0}\t{1}\t{2}\t{3}\t{4}", wknames.Value, matched.Row, matched.Column, Path.GetFileNameWithoutExtension(fileName), fileName));
                }
            }

            var rowcounter = 0;
            File.AppendAllText(@"ds\vaccines\headermatches.csv", buildr.ToString());
            foreach (var rng in availableRanges)
            {
                var rangename = rng.rangeName;
                if (!ds.ContainsKey(rangename))
                {
                    ds[rangename] = new List<string>();
                }
                var datavalues = ds[rangename];

                var rowid = rng.firstCellRowId;
                var rowmax = rng.range.Rows.Count;

                var columnid = rng.firstCellColumnId;
                var rngColmax = rng.range.Columns.Count;

                var colmax = (columnid + 15)> rngColmax? rngColmax: (columnid + 15);// maxColumnsExpected[rng.worksheetId] + columnid;
                //first row is for the headers      
                var blankRowCount = 0;
                for (var rowindx = rowid; rowindx < rowmax; rowindx++)
                {
                    //we skip the first row
                    if (rowindx == rowid)
                    {
                        //if (!addHeaderRow)
                        //    continue;
                    }

                    var drow = dt.NewRow();
                    var builder = new StringBuilder();
                    drow["id"] = rowindx;
                    drow["src_filepath"] = fileName;
                    drow["src_file"] = Path.GetFileName(fileName);
                    builder.Append(fileName + "\t" + rng.worksheetName);
                    var quitrow = false;
                    for (var colmnindx = columnid; colmnindx <= colmax; colmnindx++)
                    {
                        if (colmnindx - columnid + 3 >= fields.Count)
                        {
                            break;
                        }
                        //we check if the column is blank, if so, we break inner loop
                        var cellValue = getCellValue(rng.range, rowindx, colmnindx, "9999").checkLength2();
                        //schoolname = cellValue;
                        //we check the names
                        if (colmnindx == columnid + 1)
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
                        }
                        //we added two meta columns to this table ie. id and src_file
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
                    //if (!str.Contains("9999,9999,9999"))
                    //saveToDB();
                    datavalues.Add(builder.ToString());
                    dt.Rows.Add(drow);
                }
            }

            //we save
            dt.AcceptChanges();
            var columnList = new StringBuilder();
            dt.Columns.Cast<DataColumn>().ToList().ForEach(c => columnList.AppendFormat("{0},", c));

            db.WriteTableToDb(dt, "vaccines_list");
            return ds;
        }
    }
}
