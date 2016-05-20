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
    public class GetValuesFromReport : ExcelWorksheetReaderBase, IQueryHelper<List<DataValue>>
    {
        public List<DataValue> Execute()
        {
            return DoDataImport(SelectedProject);
        }

        public List<DataValue> DoDataImport(ProjectName projectName)
        {
            _showSimilarMessages = true;
            List<DataValue> toReturn = null;
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application() { Visible = false };
                var res = ImportData(excelApp, projectName);
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
        
        private List<DataValue> ImportData(Microsoft.Office.Interop.Excel.Application excelApp, ProjectName projectName)
        {
            PerformProgressStep("Please wait, initialising");
            var _loadAllProgramDataElements = new GetProgramAreaIndicators().GetAllProgramDataElements();

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

            //we get the facility codes
            //LocationDetail locationDetail = null;

            switch (projectName)
            {
                case ProjectName.DOD:
                    {
                        Worksheet coverWorksheet = null;
                        const string coverWorksheetName = "Cover1";
                        try
                        {
                            coverWorksheet = (Worksheet)workbook.Sheets[coverWorksheetName];
                        }
                        catch
                        {
                            ShowMissingWorksheet(coverWorksheetName);
                            return null;
                        }

                        locationDetail = GetReportLocationDetails(coverWorksheet, ShowErrorAndAbort);

                        break;
                    }
                case ProjectName.IHP_VMMC:
                    {
                        //then its VMMC, we remove the other DOD program areas and just leave one
                        if (!worksheetNames.ContainsKey(GetProgramAreaIndicators.IhpVmmcProgramAreaName))
                        {
                            throw new ArgumentNullException("Invalid file selected");
                        }

                        var vmmcProgramArea = _loadAllProgramDataElements.FirstOrDefault(t => t.ProgramArea == GetProgramAreaIndicators.dodVmmcProgramAreaName);
                        vmmcProgramArea.ProgramArea = GetProgramAreaIndicators.IhpVmmcProgramAreaName;
                        _loadAllProgramDataElements.Clear();
                        _loadAllProgramDataElements.Add(vmmcProgramArea);
                        locationDetail = GetIhpReportLocationDetails(workbook, ShowErrorAndAbort);
                        break;
                    }
                case ProjectName.IHP_Capacity_Building_and_Training:
                    {
                        break;
                    }
            }

            if (locationDetail == null)
                return null;

            var datavalues = new List<DataValue>();

            MarkStartOfMultipleSteps(_loadAllProgramDataElements.Count + 2);
            foreach (var dataElement in _loadAllProgramDataElements)
            {
                var programAreaName = dataElement.ProgramArea;
                var worksheetName = programAreaName;
                //if (isIhpVmmc && programAreaName != GetProgramAreaIndicators.dodVmmcProgramAreaName)
                //    continue;

                //if (projectName == ProjectName.IHP_VMMC)
                //{
                //    worksheetName = GetProgramAreaIndicators.IhpVmmcProgramAreaName;
                //    programAreaName = GetProgramAreaIndicators.IhpVmmcProgramAreaName;
                //}

                PerformProgressStep("Please wait, Processing worksheet: " + programAreaName);
                ResetSubProgressIndicator(dataElement.Indicators.Count + 1);

                var xlrange = ((Worksheet)workbook.Sheets[worksheetNames[worksheetName]]).UsedRange;
                //we scan the first 3 rows for the row with the age groups specified for this program area
                var rowCount = xlrange.Rows.Count;
                var colCount = xlrange.Columns.Count;

                //we have the column indexes of the first age category options, and other occurrences of the same
                var firstAgeGroupCell = GetFirstAgeGroupCell(dataElement, xlrange, ProjectName.IHP_VMMC == projectName);

                //Now we find the row indexes of the program indicators
                //ProgramIndicator currentRowMatchingIndicator = null;
                var firstIndcatorRowIndex = -1;
                for (var rowIndex = 1; rowIndex <= rowCount; rowIndex++)
                {
                    var value = getCellValue(xlrange, rowIndex, 1);
                    if (string.IsNullOrWhiteSpace(value)) continue;
                    var matchingDataElementByIndicatorId = dataElement.Indicators.FirstOrDefault(t => t.IndicatorId == value);
                    if (matchingDataElementByIndicatorId != null)
                    {
                        firstIndcatorRowIndex = rowIndex;
                        //currentRowMatchingIndicator = matchingDataElementByIndicatorId;
                        break;
                    }
                }
                //now we know that AgeGroups start from e.g. colmn 7 of row 4
                //we also know that indicators start from Row X of column 1
                //we can now start reading these values into an array for each indicator vs age group

                //we start reading the values from cell [firstIndcatorRowIndex, firstAgeGroupCell.Colmn1]
                LogCsvOutput("Processing: " + programAreaName);

                var testBuilder = new StringBuilder();
                testBuilder.AppendLine();
                testBuilder.AppendLine(programAreaName);

                var countdown = dataElement.Indicators.Count;
                var i = firstIndcatorRowIndex;

                PerformSubProgressStep();
                do
                {
                    var indicatorid = getCellValue(xlrange, i, 1);
                    if (string.IsNullOrWhiteSpace(indicatorid))
                    {
                        if (ProjectName.IHP_VMMC == projectName)
                        {
                            countdown--;
                            continue;
                        }

                        throw new ArgumentNullException(string.Format("Expected a value in Cell ( A{0}) for sheet {1}", i, programAreaName));
                    }

                    var j = firstAgeGroupCell.Column;
                    var counter = 0;

                    //or we can get the corresponding data element and see what indicators it reports under
                    while (counter < dataElement.AgeDisaggregations.Count)
                    {
                        while (counter < dataElement.AgeDisaggregations.Count)
                        {
                            var dataValue = GetDataValue(
                                xlRange: xlrange,
                                dataElement: dataElement,
                                indicatorid: indicatorid,
                                rowId: i,
                                colmnId: j,
                                counter: counter,
                                sex: getGenderText(dataElement.Gender),
                                builder: testBuilder
                                );
                            if (dataValue != null)
                                datavalues.Add(dataValue);
                            j++;
                            counter++;
                        }
                        j++;
                        counter++;
                    }

                    if (dataElement.Gender == "both")
                    {
                        j = firstAgeGroupCell.Column2;
                        counter = 0;
                        while (counter < dataElement.AgeDisaggregations.Count)
                        {
                            var dataValue = GetDataValue(
                                xlRange: xlrange,
                                dataElement: dataElement, indicatorid: indicatorid,
                                rowId: i, colmnId: j,
                                counter: counter,
                                sex: getGenderText("Female"),
                                builder: testBuilder
                                );
                            if (dataValue != null)
                                datavalues.Add(dataValue);
                            j++;
                            counter++;
                        }
                    }

                    PerformSubProgressStep();
                    testBuilder.AppendLine();
                    countdown--;
                    i++;
                } while (countdown > 0);

                LogCsvOutput(testBuilder.ToString());
                //Console.WriteLine("Done - "+ dataElement.ProgramArea);
            }

            PerformProgressStep("Please wait, finalizing");
            datavalues.ForEach(t =>
            {
                t.ReportYear = locationDetail.ReportYear;
                t.ReportMonth = locationDetail.ReportMonth;
                t.FacilityName = locationDetail.FacilityName;
            });
            PerformProgressStep("Please wait, Preparing to display results");

            //convert to dataset
            return datavalues;
        }
        
    }
}
