using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Windows;
using ZambiaDataManager.Storage;

namespace ZambiaDataManager.CodeLogic
{
    public class GetDodDataFromExcel : ExcelWorksheetReaderBase, IQueryHelper<List<DataValue>>
    {
        internal string worksheetName;
        const string CoverSheetName = "Cover1";        

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
                if (res != null)
                {
                    //add location details here
                    PerformProgressStep("Please wait, finalizing");

                    res.ForEach(t =>
                    {
                        t.ReportYear = locationDetail.ReportYear;
                        t.ReportMonth = locationDetail.ReportMonth;
                        t.FacilityName = locationDetail.FacilityName;
                    });

                    PerformProgressStep("Please wait, Preparing to display results");

                    if (IsInError)
                        return null;

                    toReturn = res;
                }
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
            //we analyse if our field dictionary has unique indicatorids
            PerformProgressStep("Please wait, initialising");
            //we have twwo spreadsheets for finance: 
            var _loadAllProgramDataElements = new
                GetProgramAreaIndicators()
                .GetDodDataElements();
            //.GetAllProgramDataElements();

            foreach (var element in _loadAllProgramDataElements)
            {
                var allIndicIds = (from t in element.Indicators
                                   select t.IndicatorId).Count();
                var distinctIndicIds = (from t in element.Indicators
                                        select t.IndicatorId).ToList().Distinct().Count();
                if (allIndicIds != distinctIndicIds)
                    throw new ArgumentOutOfRangeException(
                        "Duplicate indicatorids for program area " + element.ProgramArea);
            }

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

            if (!worksheetNames.ContainsKey(CoverSheetName.ToLowerInvariant()))
            {
                MessageBox.Show(
    string.Format(
    "Please ensure that the file \r'{0}'\r contains a Cover sheet called '{1}'", fileName, CoverSheetName)
    , "Missing worksheet", MessageBoxButton.OK);
                return null;
            }

            var coverWorksheet = (Worksheet)workbook.Sheets[worksheetNames[CoverSheetName.ToLowerInvariant()]];
            locationDetail = GetReportLocationDetails(coverWorksheet, ShowErrorAndAbort);

            var datavalues = new List<DataValue>();
            MarkStartOfMultipleSteps(_loadAllProgramDataElements.Count + 2);
            foreach (var dataElement in _loadAllProgramDataElements)
            {
                var programAreaName = dataElement.ProgramArea;
                PerformProgressStep("Please wait, Processing worksheet: " + programAreaName);
                ResetSubProgressIndicator(dataElement.Indicators.Count + 1);

                worksheetName = programAreaName.ToLowerInvariant();
                if (!worksheetNames.ContainsKey(worksheetName))
                {
                    MessageBox.Show(
                        string.Format(
                        "Please ensure that the file \r'{0}'\r contains a sheet called '{1}'", fileName, programAreaName)
                        , "Missing worksheet", MessageBoxButton.OK);

                    var choice = MessageBox.Show(string.Format("Do you still want to proceed and import the data without {0} data", programAreaName), "Proceed", MessageBoxButton.YesNo);
                    if (choice != MessageBoxResult.Yes)
                    {
                        return null;
                    }
                    else
                    {
                        continue;
                    }
                }

                var xlrange = ((Worksheet)workbook.Sheets[worksheetNames[worksheetName.ToLowerInvariant()]]).UsedRange;
                //we scan the first 3 rows for the row with the age groups specified for this program area
                var rowCount = xlrange.Rows.Count;
                var colCount = xlrange.Columns.Count;

                //we have the column indexes of the first age category options, and other occurrences of the same
                var firstAgeGroupCell = GetFirstAgeGroupCell(dataElement, xlrange, false);
                if (firstAgeGroupCell.Column == -1 && firstAgeGroupCell.Row == -1)
                {
                    MessageBox.Show(
    string.Format(
    "File \r'{0}'\r is not well formed. Please upload a file that has the correct structure", fileName)
    , "Incorrectly formatted file", MessageBoxButton.OK);
                    return null;
                }

                var ageGroupCells = GetMatchedCellsInRow(excelRange: xlrange,
                    searchTerms: dataElement.AgeDisaggregations,
                    endColumnIndex: colCount, startColumnIndex: firstAgeGroupCell.Column,
                    rowIndex: firstAgeGroupCell.Row,
                    alternateAgeLookup: null
                    );

                //Now we find the row indexes of the program indicators
                var indicatorList = (from t in dataElement.Indicators select t.Indicator).ToList();
                var firstIndicatorCell = GetFirstMatchedCellByRow(excelRange: xlrange,
                    searchTerms: indicatorList,
                    startColumnIndex: 1,
                    endColumnIndex: firstAgeGroupCell.Column - 1,
                    maxRows: rowCount,
                    statRowIndex: firstAgeGroupCell.Row + 1
                    );

                var indicatorCells = GetCellsInColumnContaining(
                    excelRange: xlrange,
                    columnIndex: firstIndicatorCell.Column,
                    searchTerms: indicatorList,
                    startRowIndex: firstIndicatorCell.Row,
                    maxRows: rowCount
                    );

                foreach (var rowObject in indicatorCells)
                {
                    //we convert the indicator to an indicatorId
                    var matchingIndicator = dataElement.Indicators.Where(t => t.Indicator == rowObject.Key).FirstOrDefault();
                    if (matchingIndicator == null)
                    {
                        throw new ArgumentOutOfRangeException("Could not match indicator " + rowObject.Key);
                    }

                    foreach (var indicatorAgeGroupCells in ageGroupCells)
                    {
                        foreach (var indicatorAgeGroupCell in indicatorAgeGroupCells.Value)
                        {
                            var dataValue = getCellValue(dataElement, matchingIndicator.IndicatorId, xlrange, rowObject, indicatorAgeGroupCells, indicatorAgeGroupCell);
                            if (dataValue == null)
                                continue;
                            datavalues.Add(dataValue);
                        }
                    }
                }
            }
            return datavalues;
        }
    }
}
