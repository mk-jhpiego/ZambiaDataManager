using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace HPVExcelReader
{
    public class GetExcelAsDataTable3 : ExcelWorksheetReaderBase, IQueryHelper<Dictionary<string, List<string>>>
    {
        public bool addHeaderRow { get; set; }
        public Microsoft.Office.Interop.Excel.Application excelApp = null;
        public bool IsInError { get; set; }
        //public DbHelper Db { get; set; }
        public List<string> expectedWorksheetNames = new List<string>() {
            "2a. Primary Schools","2b. Secondary Schools","2a.PRIMARY SCHOOLS",
            "3. Communities ","4. Vaccine and supplies and M&E"};

        public List<string> firstHeader = new List<string>() {
            "No.","Name","Village/ Township","Date of Birth","dd/mm/yyyy",
            "Grade *if in school","Age (Yrs)","Plot / House Number"};
        public List<string> preferredWorksheetname = new List<string>() {
            "Primary  School",
            "Secondary  School",
            "Community Name","Health facility" };
        public const int MAX_BLANKROWS = 10;
        public List<int> maxColumnsExpected = new List<int>() { 7, 2, 3, 3 };

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
            public string rangeName { get; internal set; }
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
        public bool fileSkipped = false;
        public bool headersMatched = true;
        public SqlConnection sqlConn;
     //   void saveToDB(Object[] savearray)
     //   {
     //       var sql = @"INSERT INTO [dbo].[hpv_register20190503]
     //      ([serial],[srcfile],[worksheet],[No]
     //      ,[Name],[Village/ Township],[Date of Birth],[grade])
     //VALUES([serial]
     //      ,[srcfile]
     //      ,[worksheet]
     //      ,[No]
     //      ,[Name]
     //      ,[Village/ Township]
     //      ,[Date of Birth]
     //      ,[grade])";
     //       using (var sqlcmd = new SqlCommand(sql, sqlConn))
     //       {
     //           if (sqlConn.State != System.Data.ConnectionState.Open)
     //           {
     //               sqlConn.Open();
     //           }
     //           sqlcmd.CommandText = "select 1666";
     //           var t = sqlcmd.ExecuteScalar();
     //           Console.WriteLine(t);
     //       }
     //   }
        private Dictionary<string, List<string>> ImportData()
        {
            var dds = new System.Data.DataSet();
            var dt = new System.Data.DataTable("hpv_list");
            dds.Tables.Add(dt);

            dt.Columns.Add("id");
            dt.Columns.Add("srcfile");
            dt.Columns.Add("worksheet");
            dt.Columns.Add("numb");
            dt.Columns.Add("names");
            dt.Columns.Add("village_township");
            dt.Columns.Add("date_of_birth");
            dt.Columns.Add("grade14");
            dt.Columns.Add("grade15");

            var unwantedSheetRows = new List<string>();
            expectedWorksheetNames.ForEach(t=> unwantedSheetRows.Add(t.Trim().ToLowerInvariant()));

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
                
                var cleanName = worksheetName.Trim().ToLowerInvariant();
                if (unwantedSheetRows.Contains(cleanName))
                {
                    fileSkipped = true;
                    break;
                }

                worksheetNames.Add(cleanName, worksheetName);
            }

            if (fileSkipped)
            {
                return null;
            }

            //we open each worksheet and search for the required columns


            
            //var sz = expectedWorksheetNames.Count();
            var workbookErrors = new StringBuilder();

            var buildr = new StringBuilder();
            var availableRanges = new List<headerIndexes>();
            var ctr = 0;
            var skipSheets = new List<string>() { "inside cover pg", "cover","instructions" };

            var includeFilenameInNoMatch = true;
            foreach (var wknames in worksheetNames)//(var ix = 0; ix < sz; ix++)
            {
                Console.WriteLine(wknames.Key);
                if(wknames.Key.Contains("cover")|| wknames.Key.Contains(" map ") || wknames.Key.Contains("instructions") || skipSheets.Contains(wknames.Key))//skipSheets.Contains(wknames.Key) || 
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
                        File.AppendAllText(@"ds\registers\nomatch.txt", string.Format("{2}{1}{0}{1}{2}{1}", fileName, Environment.NewLine, "**********"));
                    }
                    File.AppendAllText(@"ds\registers\nomatch.txt", string.Format("{2}\t{0}{1}", 
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
            File.AppendAllText(@"ds\registers\headermatches.csv", buildr.ToString());
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
                var colmax = 7;// maxColumnsExpected[rng.worksheetId] + columnid;

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
                    //var pars = new List<SqlParameter>();
                    //var pccntr = 1;
                    //var p1 = new SqlParameter("p1", fileName);
                    //pars.Add(p1);
                    //var p2 = new SqlParameter("p2", rng.worksheetName);
                    //pars.Add(p2);

                    drow[0] = rowcounter++;
                    drow[1] = fileName.checkLength();
                    drow[2] = rng.worksheetName.checkLength();
                    builder.Append(fileName + "\t" + rng.worksheetName);
                    //string schoolname = string.Empty;
                    var quitrow = false;                    
                    for(var colmnindx= columnid; colmnindx < colmax; colmnindx++)
                    {
                        //we check if the column is blank, if so, we break inner loop
                        var cellValue = getCellValue(rng.range, rowindx, colmnindx, "9999").checkLength();
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
                        drow[colmnindx - columnid + 3] = cellValue;
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
            db.WriteTableToDb(dt,"hpv_list");
            return ds;
        }
    }

}
