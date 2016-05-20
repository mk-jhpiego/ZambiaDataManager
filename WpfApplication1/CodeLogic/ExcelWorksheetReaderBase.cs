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
    public class ExcelWorksheetReaderBase
    {
        public Action<string> Alert { get; set; }
        public string fileName { get; set; }
        public LocationDetail locationDetail = null;

        public IDisplayProgress progressDisplayHelper { get; set; }
        public ProjectName SelectedProject
        {
            get;
            internal set;
        }

        protected static LocationDetail GetIhpReportLocationDetails(Workbook workbook,
            Action<string, string, string, int, int, bool> errorDialog
            )
        {
            var locationDetail = new LocationDetail();
            var xlRange = ((Worksheet)workbook.Sheets[GetProgramAreaIndicators.IhpVmmcProgramAreaName]).UsedRange;

            var facilityValue = Convert.ToString(getCellValue(xlRange, 2, 2)); //G1
            if (!string.IsNullOrWhiteSpace(facilityValue))
            {
                var split = facilityValue.Split(new[] { '>' }, StringSplitOptions.RemoveEmptyEntries);
                if (split.Length == 2)
                {
                    locationDetail.FacilityName = split[1];
                    //we validate the hmiscode
                }
                else
                {
                    throw new ArgumentException("Missing entry or value for 'Facility Name' in worksheet " + GetProgramAreaIndicators.IhpVmmcProgramAreaName);
                }
            }
            else
            {
                throw new ArgumentException("Missing entry or value for 'Facility Name' in worksheet " + GetProgramAreaIndicators.IhpVmmcProgramAreaName);
            }



            var yearValue = Convert.ToString(getCellValue(xlRange, 5, 2)); //B5
            if (!string.IsNullOrWhiteSpace(yearValue))
            {
                int year = -1;
                if (!int.TryParse(yearValue.Trim(), out year))
                {
                    errorDialog?.Invoke(yearValue, "Year", GetProgramAreaIndicators.IhpVmmcProgramAreaName, 5, 2, true);
                }
                locationDetail.ReportYear = Convert.ToInt32(yearValue.Trim());
            }
            else
            {
                throw new ArgumentException("Missing entry or value for 'Year' in worksheet " + GetProgramAreaIndicators.IhpVmmcProgramAreaName);
            }

            //report month
            var mValue = Convert.ToString(getCellValue(xlRange, 5, 6)); //F5
            var monthName = Constants.GetAlternateStandardMonthName(mValue);
            if (!string.IsNullOrWhiteSpace(monthName))
            {
                locationDetail.ReportMonth = monthName;
            }
            else
            {
                throw new ArgumentException("Missing entry or value for 'Month Reported on' in worksheet " + GetProgramAreaIndicators.IhpVmmcProgramAreaName);
            }

            return locationDetail;
        }

        protected static LocationDetail GetReportLocationDetails(Worksheet coverWorksheet, Action<string, string, string, int, int, bool> errorDialog)
        {
            var locationDetail = new LocationDetail();
            //check if not Hmiscode, month and year are specifierd. If not, we quit
            //we index the list of field headers and just get the i,j + 2 cell to haethe vakueas entered by  the user
            var coverWorksheetName = coverWorksheet.Name;
            var coverRange = coverWorksheet.UsedRange;
            var rows = coverRange.Rows.Count;
            var colmns = coverRange.Columns.Count;
            var expectedCoverFields = new List<string>() {
                "Name of Health Facility",
"Province","District","Constituency","Ward",
"Date Report Compiled","Month Reported on","Year Reported on"
};
            var coverFieldsLower = (from label in expectedCoverFields
                                    select label.ToLower()).ToList();
            var fieldIndex = new Dictionary<string, int>();

            if (colmns < 2)
                throw new ArgumentOutOfRangeException("Expected more than 2 columns for the cover page");

            var healthFacilityLabel = "Name of Health Facility".ToLowerInvariant();
            var reportMonth = "Month Reported on".ToLowerInvariant();
            var reportYear = "Year Reported on".ToLowerInvariant();

            var yearfound = false;
            for (var i = 1; i <= rows; i++)
            {
                var fieldName = Convert.ToString(getCellValue(coverRange, i, 2));
                var fieldValue = Convert.ToString(getCellValue(coverRange, i, 4));

                if (string.IsNullOrWhiteSpace(fieldName) || string.IsNullOrWhiteSpace(fieldValue))
                {
                    continue;
                }

                var lowerFieldName = fieldName.ToLowerInvariant().Trim();
                if (!coverFieldsLower.Contains(lowerFieldName))
                    continue;

                var lowerFieldValue = fieldValue.ToLowerInvariant().Trim();

                if (lowerFieldName == healthFacilityLabel)
                {
                    var split = fieldValue.Trim().Split(new[] { '>' }, StringSplitOptions.RemoveEmptyEntries);
                    if (split.Length == 2)
                    {
                        locationDetail.FacilityName = split[1];
                        //we validate the hmiscode
                    }
                    else
                    {
                        locationDetail.FacilityName = string.Empty;
                    }
                }
                else if (lowerFieldName == reportMonth)
                    locationDetail.ReportMonth = fieldValue.Trim();
                else if (lowerFieldName == reportYear)
                {
                    yearfound = true;
                    int year = -1;
                    if (!int.TryParse(fieldValue.Trim(), out year))
                    {
                        errorDialog?.Invoke(fieldValue, "Year Reported on", coverWorksheetName, i, 4, true);
                    }
                    locationDetail.ReportYear = Convert.ToInt32(fieldValue.Trim());
                }
            }

            if (string.IsNullOrWhiteSpace(locationDetail.FacilityName))
            {
                throw new ArgumentException("Missing entry or value for 'Name of Health Facility' in worksheet " + coverWorksheetName);
            }
            if (!yearfound)
            {
                throw new ArgumentException("Missing entry or value for 'Year Reported on' in worksheet " + coverWorksheetName);
            }

            if (string.IsNullOrWhiteSpace(locationDetail.ReportMonth))
            {
                throw new ArgumentException("Missing entry or value for 'Month Reported on' in worksheet " + coverWorksheetName);
            }
            else
            {
                var monthName = Constants.GetStandardMonthName(locationDetail.ReportMonth);
                locationDetail.ReportMonth = monthName;
            }
            return locationDetail;
        }



        public void PerformProgressStep(string message = "")
        {
            progressDisplayHelper.PerformProgressStep(message);
        }

        public void MarkStartOfMultipleSteps(int stepsToExpect)
        {
            progressDisplayHelper.MarkStartOfMultipleSteps(stepsToExpect);
        }

        public void ResetSubProgressIndicator(int stepsToExpect)
        {
            progressDisplayHelper.ResetSubProgressIndicator(stepsToExpect);
        }

        public void PerformSubProgressStep()
        {
            progressDisplayHelper.PerformSubProgressStep();
        }

        protected void LogCsvOutput(string text)
        {
            //File.AppendAllText("valuesRead.csv", text);
        }

        protected string getGenderText(string genderName)
        {
            var t = genderName == "both" ? "Male" :
                                (genderName == "Male" || genderName == "Female" ? genderName : "");
            return t;
        }

        protected object initProgDatElmts = new object();

        protected static List<string> customHandledIndicators = new List<string>() { "FP4", "FP6", "FP7" };
        protected DataValue GetDataValue(Range xlRange, ProgramAreaDefinition dataElement, string indicatorid, int rowId, int colmnId, int counter, string sex, StringBuilder builder = null)
        {
            var i = rowId;
            var j = colmnId;
            var value = getCellValue(xlRange, i, j);
            double asDouble;
            DataValue dataValue = null;

            if (customHandledIndicators.Contains(indicatorid))
                return null;

            try
            {
                asDouble = value.ToDouble();
                if (asDouble == -2146826273 || asDouble == -2146826281)
                {
                    ShowErrorAndAbort(value, indicatorid, dataElement.ProgramArea, i, j);
                    return null;
                }
            }
            catch
            {
                ShowErrorAndAbort(value, indicatorid, dataElement.ProgramArea, i, j);
                return null;
            }

            if (asDouble != Constants.NOVALUE)
            {
                if (value == null)
                {
                    ShowValueNullErrorAndAbort(indicatorid, dataElement.ProgramArea, i, j);
                }

                dataValue = new DataValue()
                {
                    IndicatorValue = asDouble,
                    IndicatorId = indicatorid,
                    ProgramArea = dataElement.ProgramArea,
                    AgeGroup = dataElement.AgeDisaggregations[counter],
                    Sex = sex
                };
                if (builder != null)
                {
                    LogCsvOutput(string.Format("{0}\t", value));
                }
            }
            else
            {
                if (builder != null)
                {
                    LogCsvOutput(string.Format("{0}\t", "x"));
                    //builder.AppendFormat("{0}\t", "x");
                }
            }
            return dataValue;
        }

        protected void ShowMissingWorksheet(string coverWorksheetName)
        {
            IsInError = true;
            MessageBox.Show("Error trying to access the worksheet " + coverWorksheetName);
        }

        protected static string GetColumnName(int index)
        {
            return (index > 26 ? "A" : "") + (index == 0 ? 'A' : Convert.ToChar('A' + index % 26 - 1)).ToString();
        }
        protected bool IsInError = false;
        protected static bool _showSimilarMessages = true;
        protected void ShowErrorAndAbort(string value, string indicatorid, string programArea, int i, int j, bool throwException = false)
        {
            IsInError = true;
            if (!_showSimilarMessages) return;

            if (throwException)
            {
                throw new ArgumentException(string.Format("Could not convert value '{0}' in worksheet '{1}' and Cell ({3}{2}) as a number", value, programArea, i, GetColumnName(j)));
            };
            //var dialog = new Microsoft.Win32.CommonDialog();

            var res = MessageBox.Show(string.Format("Could not convert value '{0}' in worksheet '{1}' and Cell ({3}{2}) as a number. \nDo you want to see other similar messages", value, programArea, i, GetColumnName(j)),
                "Error getting value. The tool will quit.", MessageBoxButton.YesNo

                //MessageBoxButtons.YesNo

                );
            if (res != MessageBoxResult.Yes)
            {
                _showSimilarMessages = false;
            }
        }

        protected void ShowValueNullErrorAndAbort(string indicatorid, string programArea, int i, int j)
        {
            IsInError = true;
            if (!_showSimilarMessages) return;

            var res = MessageBox.Show(string.Format("Could not determine the value in worksheet '{0}' and Cell ({2}{1}). Check that the cells are not merged", programArea, i, GetColumnName(j)), "Error getting value. The tool will quit.", MessageBoxButton.YesNo);
            if (res != MessageBoxResult.Yes)
            {
                _showSimilarMessages = false;
            }
        }

        protected static Dictionary<string, RowColmnPair> GetCellsInColumnContaining(Range excelRange, int columnIndex, List<string> searchTerms,
            int startRowIndex, int maxRows)
        {
            var indicatorCells = new Dictionary<string, RowColmnPair>();
            for (var rowIndex = startRowIndex; rowIndex <= maxRows; rowIndex++)
            {
                var value = getCellValue(excelRange, rowIndex, columnIndex);
                if (string.IsNullOrWhiteSpace(value)) continue;
                if (searchTerms.Contains(value))
                {
                    indicatorCells[value] = new RowColmnPair() { Column = columnIndex, Row = rowIndex };
                }
            }
            return indicatorCells;
        }

        protected static Dictionary<string, RowColmnPair> GetMatchedCellsInRow(Range excelRange, List<string> searchTerms,
            int rowIndex, int startColumnIndex, int endColumnIndex)
        {
            var toReturn = new Dictionary<string, RowColmnPair>();
            for (var colmnId = startColumnIndex; colmnId <= endColumnIndex; colmnId++)
            {
                var value = getCellValue(excelRange, rowIndex, colmnId);

                //people might have other columns, so we pick what we want
                if (string.IsNullOrWhiteSpace(value) || value.Length > 40) continue;

                //we find the first match, and quit once we read all columns in that row matching disaggregations
                if (searchTerms.Contains(value))
                {
                    //we get all locations for our desired  row
                    toReturn[value] = new RowColmnPair() { Column = colmnId, Row = rowIndex };
                }
            }
            return toReturn;
        }

        protected static RowColmnPair GetFirstMatchedCellByRow(Range excelRange, List<string> searchTerms, int statRowIndex, int maxRows, int startColumnIndex, int endColumnIndex)
        {
            RowColmnPair firstIndicatorCell = null;
            for (var rowIndex = statRowIndex; rowIndex <= maxRows; rowIndex++)
            {
                for (var colIndex = endColumnIndex; colIndex >= startColumnIndex; colIndex--)
                {
                    var value = getCellValue(excelRange, rowIndex, colIndex);
                    if (string.IsNullOrWhiteSpace(value)) continue;
                    if (searchTerms.Contains(value))
                    {
                        firstIndicatorCell = new RowColmnPair()
                        {
                            Column = colIndex,
                            Row = rowIndex
                        };
                        break;
                    }
                }
                if (firstIndicatorCell != null) break;
            }
            return firstIndicatorCell;
        }

        protected static RowColmnPair GetFirstAgeGroupCell(ProgramAreaDefinition dataElement, Range xlrange, bool isNonDod)
        {
            int colCount = xlrange.Columns.Count;
            colCount = colCount > 15 ? 12 : colCount;

            int row = -1, colmn = -1, colmn2 = -1;

            var matchfound = false;
            var maxDepthSearchRows = 8;
            //var firstAgeGroupColumn = -1;
            for (var rowId = 1; rowId <= maxDepthSearchRows; rowId++)
            {
                for (var colmnId = 1; colmnId <= colCount; colmnId++)
                {
                    var value = getCellValue(xlrange, rowId, colmnId);
                    if (string.IsNullOrWhiteSpace(value) || value.Length > 40) continue;

                    if (dataElement.AgeDisaggregations.Contains(value))
                    {
                        //we've found our column, time to find where the data begibs, lets find the corresponding indicator
                        //we'll scan for columns from rowid to perhaps 5 places, and starting from column 0
                        row = rowId;
                        colmn = colmnId;
                        matchfound = true;

                        if (dataElement.Gender.ToLowerInvariant() == "both")
                        {
                            //we continue and find the next occurrence of this value                            
                            colmn2 = findNextOccurence(dataElement, xlrange, colCount, rowId, colmnId + 1, value);
                        }

                        break;
                    }
                }
                if (matchfound) break;
            }

            return new RowColmnPair(row, colmn, colmn2);
        }

        public static string getCellValue(Range xlrange, int rowId, int colmnId)
        {
            var cellvalue = Convert.ToString(xlrange[rowId, colmnId].Value);
            return cellvalue == null ? string.Empty : cellvalue.ToString().Trim();
        }

        protected static int findNextOccurence(ProgramAreaDefinition dataElement, Range xlrange, int colCount, int rowId, int startColmnIndex, string valueToFind)
        {
            //if(dataElement.ProgramArea =="PEP")
            int colmnIndex = -1;
            for (var colmnId = startColmnIndex; colmnId <= colCount; colmnId++)
            {
                var value = getCellValue(xlrange, rowId, colmnId);
                if (value != valueToFind)
                    continue;
                colmnIndex = colmnId;
                break;
            }
            return colmnIndex;
        }
    }


}
