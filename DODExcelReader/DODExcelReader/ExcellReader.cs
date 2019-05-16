using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace DODExcelReader
{
    public class GetExcelAsDataTable2 : ExcelWorksheetReaderBase, IQueryHelper<Dictionary<string, List<string>>>
    {
        public bool addHeaderRow { get; set; }
        public Microsoft.Office.Interop.Excel.Application excelApp = null;
        public bool IsInError { get; set; }
        public string rootPath { get; internal set; }

        //public DbHelper Db { get; set; }
        public static List<string> expectedWorksheetNames = new List<string>() { "FacilityData" };
        public static List<string> firstHeader = new List<string>() {
            "FacilityID" };
        public static List<int> maxColumnsExpected = new List<int>() { 7};
        public static List<string> expectedColumns = new List<string>() { "FacilityID", "IndicatorID", "ReferenceYear", "ReferenceMonth", "Sex", "AgeGroup", "Number" };
        //"FacilityID","IndicatorID","ReferenceYear","ReferenceMonth","Sex","AgeGroup","Number"

        public DataSet ExecuteDataset()
        {
            //ProjectName projectName
            DataSet toReturn = null;
            try
            {
                var res = ImportDataset();
                //add location details here
                if (IsInError)
                    return null;
                toReturn = res;
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("{0} --ERROR-- {1}", fileName, ex.Message));
                File.AppendAllText(@"ds\importfail.txt", fileName);
                File.AppendAllText(@"ds\importfail_errors.txt", fileName + "\r\n" + ex.Message + "\r\n");
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
            return cellvalue == null || cellvalue.ToLowerInvariant() == "null" ? (valueIfNull == "" ? string.Empty : valueIfNull) : cellvalue.ToString().Trim();
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

        private DataSet ImportDataset()
        {
            var ds = new Dictionary<string, List<string>>();
            //var dataObjects = new List<dataObject>();
            //var fileShortName = Path.GetFileNameWithoutExtension(fileName);
            var dir = Path.GetDirectoryName(fileName);
            dir = dir.Replace(rootPath, "");
            Console.WriteLine(dir);
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
                    workbookErrors.AppendLine("\t" + name);
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
                        range = xlrange
                    });
                    Console.WriteLine(string.Format("{2},\t{0},\t{1}", matched.Row, matched.Column, name));
                    buildr.AppendLine(string.Format("{0},{1},{2},{3},{4}", name, matched.Row, matched.Column, Path.GetFileNameWithoutExtension(fileName), fileName));

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
            var dset = new DataSet();
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
                //StagingTable
                var rndid = new Random(DateTime.Now.Millisecond).Next(1000, 10000);
                var tempname = string.Format("temp_facdata_{0}", rndid);
                var dtable = new System.Data.DataTable(tempname);
                //FacilityID	IndicatorID	ReferenceYear	ReferenceMonth	Sex	AgeGroup	Number
                dtable.Columns.AddRange(new List<DataColumn>() {
                    new DataColumn{ ColumnName="dirpath",DataType=typeof(String)},
                    new DataColumn{ ColumnName="srcfile",DataType=typeof(String)},
                    new DataColumn{ ColumnName="FacilityID",DataType=typeof(int)},
                    new DataColumn{ ColumnName="IndicatorID",DataType=typeof(String)},
                    new DataColumn{ ColumnName="ReferenceYear",DataType=typeof(int)},
                    new DataColumn{ ColumnName="ReferenceMonth",DataType=typeof(int)},
                    new DataColumn{ ColumnName="Sex",DataType=typeof(int)},
                    new DataColumn{ ColumnName="AgeGroup",DataType=typeof(int)},
                    new DataColumn{ ColumnName="Number",DataType=typeof(int)},
                    //new DataColumn{ ColumnName="dirpath",DataType=typeof(string)},
                }.ToArray());
                dset.Tables.Add(dtable);
                dtable.AcceptChanges();

                for (var rowindx = rowid; rowindx < rowmax; rowindx++)
                {
                    //we skip the first row
                    if (rowindx == rowid)
                    {
                        if (!addHeaderRow)
                            continue;
                    }
                    var drow = dtable.NewRow();
                    var builder = new StringBuilder();
                    //builder.Append(Path.GetFileNameWithoutExtension(fileName));
                    drow[0] = dir;
                    drow[1] = Path.GetFileNameWithoutExtension(fileName);
                    builder.AppendFormat("{0},{1}", dir, Path.GetFileNameWithoutExtension(fileName));
                    string schoolname = string.Empty;
                    var quitrow = false;
                    for (var colmnindx = columnid; colmnindx < colmax; colmnindx++)
                    {
                        var columnName = expectedColumns[colmnindx - columnid];
                        //we check if the column is blank, if so, we break inner loop
                        var cellValue = getCellValue(rng.range, rowindx, colmnindx, "9999").Replace(",", "-").Replace("\"", "-");
                        //cellValue = cellValue.ToLowerInvariant() == "null" ? "9999" : cellValue;
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

                        if("IndicatorID"== columnName)
                        {
                            drow[columnName] = cellValue;
                        }
                        else
                        {
                            if (string.IsNullOrWhiteSpace(cellValue))
                            {
                                cellValue = "0";
                                Console.WriteLine(string.Format("Changing to '0'"));
                            }

                            var pint = -1;
                            if(int.TryParse(cellValue, out pint))
                            {
                                drow[columnName] = pint;
                            }
                            else
                            {
                                Console.WriteLine(string.Format("Error converting to int: {0}"));
                            }
                            
                        }
                        
                        builder.AppendFormat(",{0}", cellValue);
                    }
                    if (quitrow)
                    {
                        //no need to add
                        break;
                    }
                    var str = builder.ToString();
                    if (!str.Contains("9999,9999,9999"))
                    {
                        if (rowindx % 213 == 0)
                            Console.WriteLine("{0}/{1}", rowindx, rowmax);

                        dtable.Rows.Add(drow);
                        datavalues.Add(builder.ToString());
                    }
                }

            }
            return dset;
        }

        private Dictionary<string, List<string>> ImportData()
        {
            var ds = new Dictionary<string, List<string>>();
            //var dataObjects = new List<dataObject>();
            //var fileShortName = Path.GetFileNameWithoutExtension(fileName);
            var dir = Path.GetDirectoryName(fileName);
            dir = dir.Replace(rootPath, "");
            Console.WriteLine(dir);
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
                    Console.WriteLine(string.Format("{2},\t{0},\t{1}", matched.Row, matched.Column, name));
                    buildr.AppendLine(string.Format("{0},{1},{2},{3},{4}", name, matched.Row, matched.Column, Path.GetFileNameWithoutExtension(fileName), fileName));

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
            var dset = new DataSet();
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
                
                var dtable = new System.Data.DataTable(rangename);
                //FacilityID	IndicatorID	ReferenceYear	ReferenceMonth	Sex	AgeGroup	Number
                dtable.Columns.AddRange(new List<DataColumn>() {
                    new DataColumn{ ColumnName="dirpath",DataType=typeof(String)},
                    new DataColumn{ ColumnName="srcfile",DataType=typeof(String)},
                    new DataColumn{ ColumnName="FacilityID",DataType=typeof(int)},
                    new DataColumn{ ColumnName="IndicatorID",DataType=typeof(String)},
                    new DataColumn{ ColumnName="ReferenceYear",DataType=typeof(int)},
                    new DataColumn{ ColumnName="ReferenceMonth",DataType=typeof(int)},
                    new DataColumn{ ColumnName="Sex",DataType=typeof(int)},
                    new DataColumn{ ColumnName="AgeGroup",DataType=typeof(int)},
                    new DataColumn{ ColumnName="Number",DataType=typeof(int)},
                    //new DataColumn{ ColumnName="dirpath",DataType=typeof(string)},
                }.ToArray());
                dset.Tables.Add(dtable);
                dtable.AcceptChanges();

                for (var rowindx = rowid; rowindx <= rowmax; rowindx++)
                {
                    //we skip the first row
                    if (rowindx == rowid)
                    {
                        if (!addHeaderRow)
                            continue;
                    }

                    //if (rowindx >= 10)
                    //{
                    //    break;
                    //}

                    var drow = dtable.NewRow();
                    var builder = new StringBuilder();
                    //builder.Append(Path.GetFileNameWithoutExtension(fileName));
                    drow[0] = dir;
                    drow[1] = Path.GetFileNameWithoutExtension(fileName);
                    builder.AppendFormat("{0},{1}", dir, Path.GetFileNameWithoutExtension(fileName));
                    string schoolname = string.Empty;
                    var quitrow = false;
                    for(var colmnindx= columnid; colmnindx < colmax; colmnindx++)
                    {
                        var columnName = expectedColumns[colmnindx - columnid];
                        //we check if the column is blank, if so, we break inner loop
                        var cellValue = getCellValue(rng.range, rowindx, colmnindx, "9999").Replace(",", "-").Replace("\"", "-");
                        //cellValue = cellValue.ToLowerInvariant() == "null" ? "9999" : cellValue;
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
                        drow[columnName] = cellValue;
                        builder.AppendFormat(",{0}", cellValue);
                    }
                    if (quitrow)
                    {
                        //no need to add
                        break;
                    }
                    var str = builder.ToString();
                    if (!str.Contains("9999,9999,9999"))
                    {
                        dtable.Rows.Add(drow);
                        datavalues.Add(builder.ToString());
                    }
                }
                
            }

            //return dset;
            return ds;
        }
    }

}
