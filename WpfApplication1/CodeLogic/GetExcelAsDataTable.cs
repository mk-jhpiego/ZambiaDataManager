//using Excel = Microsoft.Office.Interop.Excel;

using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
//using System.Data;
using System.Windows;

namespace ZambiaDataManager.CodeLogic
{
    public class GetFinanceDataFromExcel : ExcelWorksheetReaderBase, IQueryHelper<List<DataValue>>
    {
        internal string worksheetName;

        public GetFinanceDataFromExcel()
        {

        }

        public List<DataValue> Execute()
        {
            return DoDataImport();
        }

        public List<DataValue> DoDataImport()
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
                worksheetNames.Add(worksheetName.Trim(), worksheetName);
            }

            var datavalues = new List<DataValue>();

            MarkStartOfMultipleSteps(_loadAllProgramDataElements.Count + 2);
            foreach (var dataElement in _loadAllProgramDataElements)
            {
                var programAreaName = dataElement.ProgramArea;
                //var worksheetName = programAreaName;
                //if (isIhpVmmc && programAreaName != GetProgramAreaIndicators.dodVmmcProgramAreaName)
                //    continue;

                //if (projectName == ProjectName.IHP_VMMC)
                //{
                //    worksheetName = GetProgramAreaIndicators.IhpVmmcProgramAreaName;
                //    programAreaName = GetProgramAreaIndicators.IhpVmmcProgramAreaName;
                //}

                PerformProgressStep("Please wait, Processing worksheet: " + programAreaName);
                ResetSubProgressIndicator(dataElement.Indicators.Count + 1);

                //we get the first value if only one sheet
                if (string.IsNullOrWhiteSpace(worksheetName)&& worksheetNames.Count ==1)
                {
                    worksheetName = worksheetNames.Values.FirstOrDefault();
                }

                var xlrange = ((Worksheet)workbook.Sheets[worksheetNames[worksheetName]]).UsedRange;
                //we scan the first 3 rows for the row with the age groups specified for this program area
                var rowCount = xlrange.Rows.Count;
                var colCount = xlrange.Columns.Count;

                //we have the column indexes of the first age category options, and other occurrences of the same
                var firstAgeGroupCell = GetFirstAgeGroupCell(dataElement, xlrange, false);
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
                    foreach (var columnObject in ageGroupCells)
                    {

                        var value = getCellValue(xlrange, rowObject.Value.Row, columnObject.Value.Column);
                        var asDouble = 0d;
                        try
                        {
                            asDouble = value.ToDouble();
                            if (asDouble == 0)
                                continue;

                            if (asDouble == -2146826273 || asDouble == -2146826281)
                            {
                                ShowErrorAndAbort(value, rowObject.Key, dataElement.ProgramArea, rowObject.Value.Row, columnObject.Value.Column);
                                //return null;
                            }
                        }
                        catch
                        {
                            ShowErrorAndAbort(value, rowObject.Key, dataElement.ProgramArea, rowObject.Value.Row, columnObject.Value.Column);
                            //return null;
                        }

                        if (asDouble != Constants.NOVALUE)
                        {
                            var dataValue = new DataValue()
                            {
                                IndicatorValue = asDouble,
                                IndicatorId = rowObject.Key,
                                ProgramArea = dataElement.ProgramArea,
                                AgeGroup = columnObject.Key,
                                Sex = "Male",

                            };
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