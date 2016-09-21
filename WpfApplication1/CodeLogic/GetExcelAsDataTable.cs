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
    public class GetFinanceDataFromExcel : ExcelWorksheetReaderBase, IQueryHelper<List<DataValue>>
    {
        internal string worksheetName;

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
            PerformProgressStep("Please wait, initialising");
            //we have twwo spreadsheets for finance: 
            var _loadAllProgramDataElements = new GetProgramAreaIndicators().GetFinanceDataElements();

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

            if (!worksheetNames.ContainsKey(worksheetName.ToLowerInvariant()))
            {
                MessageBox.Show(
                    string.Format(
                    "Please ensure that the file \r'{0}'\r contains a sheet called '{1}'", fileName, worksheetName)
                    , "Missing worksheet", MessageBoxButton.OK);
                return null;
            }

            var datavalues = new List<DataValue>();

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
                    rowIndex: firstAgeGroupCell.Row
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

            PerformProgressStep("Please wait, finalizing");
            return datavalues;
        }
    }
}