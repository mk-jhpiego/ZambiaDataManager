using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ZambiaDataManager.CodeLogic
{
    public class GetExcelAsDataTable2 : ExcelWorksheetReaderBase, IQueryHelper<List<DataValue>>
    {
        internal string worksheetName;
        public int reportYear { get; set; }
        public string reportMonth { get; set; }

        public List<DataValue> Execute()
        {
            //ProjectName projectName
            _showSimilarMessages = true;
            List<DataValue> toReturn = null;
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application() { Visible = false };
                var res = ImportData(excelApp);
                //add location details here

                if (IsInError)
                    return null;
                toReturn = res;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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

        private List<DataValue> ImportData(Microsoft.Office.Interop.Excel.Application excelApp)
        {
            var acts = new List<string>();
            acts.Add(string.Format("Step 1 - {0}", DateTime.Now));
            PerformProgressStep("Please wait, initialising");
            //we have twwo spreadsheets for finance: 
            var _loadAllProgramDataElements = new GetProgramAreaIndicators().GetFinanceDataElements("staticdata//timesheet.json");

            PerformProgressStep("Please wait, Opening Excel document");

            //for all available program areas, we read the values into an array
            var workbook = excelApp.Workbooks.Open(fileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            //we get the other data
            var worksheetCount = workbook.Sheets.Count;
            var worksheetNames = new Dictionary<string, string>();
            for (var indx = 1; indx <= worksheetCount; indx++)
            {
                var worksheetName = ((Worksheet)(workbook.Sheets[indx])).Name;
                worksheetNames.Add(worksheetName.Trim().ToLowerInvariant(), worksheetName);
            }
            acts.Add(string.Format("Step 2 - {0}", DateTime.Now));

            if (!worksheetNames.ContainsKey(worksheetName.ToLowerInvariant()))
            {
                MessageBox.Show(
                    string.Format(
                    "Please ensure that the file \r'{0}'\r contains a sheet called '{1}'", fileName, worksheetName)
                    , "Missing worksheet", MessageBoxButton.OK);
                return null;
            }

            var datavalues = new List<DataValue>();
            acts.Add(string.Format("Step 3 - {0}", DateTime.Now));
            MarkStartOfMultipleSteps(_loadAllProgramDataElements.Count + 2);
            foreach (var dataElement in _loadAllProgramDataElements)
            {
                
                var programAreaName = dataElement.ProgramArea;
                PerformProgressStep("Please wait, Processing worksheet: " + programAreaName);
                ResetSubProgressIndicator(dataElement.Indicators.Count + 1);

                var xlrange = ((Worksheet)workbook.Sheets[worksheetNames[worksheetName.ToLowerInvariant()]]).UsedRange;
                //we scan the first 3 rows for the row with the age groups specified for this program area
                var rowCount = xlrange.Rows.Count;
                var colCount = xlrange.Columns.Count;

                //  xlrange.Rows

                //we have the column indexes of the first age category options, and other occurrences of the same
                var firstAgeGroupCell = GetFirstAgeGroupCell(dataElement, xlrange);
                if (firstAgeGroupCell.Column == -1 && firstAgeGroupCell.Row == -1)
                {
                    MessageBox.Show(
    string.Format(
    "File \r'{0}'\r is not well formed. Please upload a file that has the correct structure", fileName)
    , "Incorrectly formatted file", MessageBoxButton.OK);
                    return null;
                }
                acts.Add(string.Format("Step 4 - {0}", DateTime.Now));
                var ageGroupCells = GetMatchedCellsInRow(excelRange: xlrange,
                    searchTerms: dataElement.AgeDisaggregations,
                    endColumnIndex: colCount, startColumnIndex: firstAgeGroupCell.Column,
                    rowIndex: firstAgeGroupCell.Row,
                    alternateAgeLookup: dataElement.AgeDisaggregations.ToDictionary(t => t.toCleanAge(), v => v)
                    );

                acts.Add(string.Format("Step 5 - {0}", DateTime.Now));
                //Now we find the row indexes of the program indicators
                var indicatorList = (from t in dataElement.Indicators select t.Indicator).ToList();
                var firstIndicatorCell = GetFirstMatchedCellByRow(excelRange: xlrange,
                    searchTerms: indicatorList,
                    startColumnIndex: 1,
                    endColumnIndex: firstAgeGroupCell.Column - 1,
                    maxRows: rowCount,
                    statRowIndex: firstAgeGroupCell.Row + 2
                    );

                acts.Add(string.Format("Step 6 - {0}", DateTime.Now));
                var indicatorCells = GetCellsInColumnContaining(
                    excelRange: xlrange,
                    columnIndex: firstIndicatorCell.Column,
                    searchTerms: indicatorList,
                    startRowIndex: firstIndicatorCell.Row,
                    maxRows: rowCount
                    );

                acts.Add(string.Format("Step 7 Start Rows - {0}", DateTime.Now));
                var rcount = 0;
                foreach (var rowObject in indicatorCells)
                {
                    
                    var matchingIndicator = dataElement.Indicators.Where(t => t.Indicator == rowObject.Key).FirstOrDefault();
                    if (matchingIndicator == null)
                    {
                        throw new ArgumentOutOfRangeException("Could not match indicator " + rowObject.Key);
                    }

                    foreach (var indicatorAgeGroupCells in ageGroupCells)
                    {
                        foreach (var indicatorAgeGroupCell in indicatorAgeGroupCells.Value)
                        {
                            var dataValue = getCellValue(dataElement, 
                                matchingIndicator.IndicatorId, xlrange, rowObject, 
                                indicatorAgeGroupCells, indicatorAgeGroupCell);
                            if (dataValue == null)
                                continue;

                            //people are facilities and indicator is 'Project LOE'
                            var hoursCharged = new DataValue()
                            {
                                FacilityName = dataValue.IndicatorId,

                                IndicatorValue = dataValue.IndicatorValue,
                                IndicatorId = "Project LOE",
                                ProgramArea = dataValue.ProgramArea,
                                AgeGroup = dataValue.AgeGroup,
                                ReportYear = reportYear,
                                ReportMonth = reportMonth,
                                Sex= dataValue.Sex
                            };

                            datavalues.Add(hoursCharged);
                        }
                    }
                    rcount++;
                    acts.Add(string.Format("Step 8: Row {0} Processed - {1}", rcount, DateTime.Now));
                }
            }

            var g = "";
            acts.ForEach(t => g += t);

            PerformProgressStep("Please wait, finalizing");
            return datavalues;
        }
    }
}