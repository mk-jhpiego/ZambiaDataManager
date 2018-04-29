using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ZambiaDataManager.CodeLogic;
using ZambiaDataManager.Popups;
using ZambiaDataManager.Storage;
using ZambiaDataManager.Utilities;

namespace ZambiaDataManager.Forms
{
    /// <summary>
    /// Interaction logic for TimesheetImporter.xaml
    /// </summary>
    public partial class TimesheetImporter : Page
    {
        public TimesheetImporter()
        {
            InitializeComponent();
            SelectedFiles = new List<FinanceDetails>();
        }

        public ProjectName CurrentProjectName
        {
            get;
            internal set;
        }

        //ProjectName CurrentProjectName = ProjectName.General;
        public List<FinanceDetails> SelectedFiles { get; private set; }
        //public List<MatchedDataValue> ProcessedDetails { get; private set; }

        private void selectFile(object sender, RoutedEventArgs e)
        {
            SelectedFiles.Clear();

            var dialog = new Microsoft.Win32.OpenFileDialog()
            {
                CheckFileExists = true,
                Multiselect = false,
                CheckPathExists = true,
                Filter = "Excel (*.xlsx)|*.xlsx|Excel (*.xls)|*.xls",
                Title = "Please select the files to import"
            };
            var dialogResult = dialog.ShowDialog() ?? false;
            if (dialogResult)
            {
                var filename = dialog.FileNames.FirstOrDefault();
                var reportYearAndMonth = validateFileName(filename);
                if (reportYearAndMonth == null)
                {
                    MessageBox.Show("Couldn't determine the Year and Month for the files selected. Please ensure the file is named e.g. 'LOE allocation Report March 2018'");
                    return;
                }
                else
                {
                    ////we assume 0 and 1 as the month and year
                    //MessageBox.Show(
                    //    string.Format("Month and Year deduced from filename is: Month: {0}, Year: {1}", 
                    //    reportYearAndMonth.ReportMonth, reportYearAndMonth.ReportYear));
                    ////return;
                }
                //we process for report year and month
                SelectedFiles.Add(new FinanceDetails()
                {
                    FileName = filename,
                    ReportYear = reportYearAndMonth.ReportYear,
                    ReportMonth = reportYearAndMonth.ReportMonth
                });
            }

            //we refresh the grid
            //if (skip) return;

            gIntermediateData.Visibility = Visibility.Collapsed;
            if (gSelectedFiles.Height == 0)
            {
                gSelectedFiles.Height = defaultGridSize;
            }
            gSelectedFiles.ItemsSource = "";
            gSelectedFiles.ItemsSource = SelectedFiles;
            gSelectedFiles.Visibility = Visibility.Visible;

            //update the tip
            tHelpfulTip.Content = "Click Select Files to add more files or select next option";
        }

        //bool skip = true;
        void EnableSaveButtons(bool viewState) { }

        private void ShowGridDisplayPort(List<DataValue> table1, List<DataValue> table2)
        {
            //refreshDataGrid(false);
            gIntermediateData.ItemsSource = "";
            gIntermediateData.ItemsSource = table1;
            if (gIntermediateData.Height == 0)
            {
                gIntermediateData.Height = defaultGridSize;
            }
        }

        private void ResetAllGrids()
        {
            //if (skip) return;

            gSelectedFiles.Height = 0;
            gIntermediateData.Height = 0;
            gSelectedFiles.ItemsSource = "";
            gIntermediateData.ItemsSource = "";
            if (gIntermediateData.Height == 0)
            {
                gIntermediateData.Height = defaultGridSize;
            }
            SelectedFiles.Clear();
            ExcelDataValues.Clear();
        }

        double defaultGridSize = 500;
        private void reviewSelectedFiles(object sender, RoutedEventArgs e)
        {
            if (SelectedFiles.Count == 0)
                return;

            //fire up template reader and let it do the rest
            var selectedFiles = SelectedFiles;
            var res = ReadDataFiles(SelectedFiles, CurrentProjectName);
            if (res)
            {
                SelectedFiles.Clear();
            }
        }

        CodeRunner<List<DataValue>> _runner;
        public volatile List<DataValue> ExcelDataValues = null;

        public class FilesSelectedInfo
        {
            public string OfficeAllocationFile;
            public string TotalCostsFile;

            public LocationDetail LocationDetails;
        }

        public FilesSelectedInfo AssignSelectedFiles(List<FinanceDetails> files)
        {
            var thisYear = DateTime.Now.Year;
            var officeAllocationFile = string.Empty;
            var totalCostsFile = string.Empty;
            for (var i = 0; i < 2; i++)
            {
                var filename = files[i].FileName;
                var isOfficeAllocationFile = filename.ToLowerInvariant().Contains(Constants.OFFICE_ALLOCATION);
                if (isOfficeAllocationFile && !string.IsNullOrWhiteSpace(officeAllocationFile))
                {
                    //means all are office allocation files
                    MessageBox.Show("More than one Office Allocation file selected.");
                    return null;
                }
                if (i > 0 && !isOfficeAllocationFile && string.IsNullOrWhiteSpace(officeAllocationFile))
                {
                    //means no office allocation file selected
                    MessageBox.Show("Please ensure that one Office Allocation file is selected.");
                    return null;
                }
                if (isOfficeAllocationFile)
                {
                    officeAllocationFile = filename;
                }
                else
                {
                    totalCostsFile = filename;
                }
            }

            var singlefile = files.FirstOrDefault() as FinanceDetails;
            return new FilesSelectedInfo()
            {
                OfficeAllocationFile = officeAllocationFile,
                LocationDetails = new LocationDetail() { ReportMonth = singlefile.ReportMonth, ReportYear = singlefile.ReportYear },
                TotalCostsFile = totalCostsFile
            };
        }

        LocationDetail validateFileName(string selectedFileName)
        {
            var selectedFileNameNoExt = System.IO.Path.GetFileNameWithoutExtension(selectedFileName);
            //expected name is of the form 
            //LOE allocation Report March 2018 (003).xlsx
            var simplestname = "LOE allocation Report".ToLowerInvariant();
            //we'll have month space year space other stuff
            var month_year_part = selectedFileNameNoExt.ToLowerInvariant()
                .Replace(simplestname, "");
            return GetReportYearAndMonthFromFileNames(month_year_part, DateTime.Now.Year);
        }

        bool ReadDataFiles(List<FinanceDetails> files, ProjectName projectName)
        {
            if (files.Count != 1)
                throw new ArgumentOutOfRangeException("Expected one file passed in");
            var singlefile = files.FirstOrDefault();
            //processedFilesInfo.LocationDetails.FacilityName = "5040HQ5";
            if (ExcelDataValues == null) { ExcelDataValues = new List<DataValue>(); }
            else { ExcelDataValues.Clear(); }
            var officeAllocationFileData = new GetExcelAsDataTable2()
            {
                reportYear = singlefile.ReportYear,
                reportMonth = singlefile.ReportMonth,
                locationDetail = null,

                fileName = singlefile.FileName,
                worksheetName = "Sheet1",
                progressDisplayHelper = new WaitDialog()
                {
                    WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner
                },
                SelectedProject = ProjectName.General
            }.Execute();

            if (officeAllocationFileData == null)
            {
                //we skip and alert the user of the error
                return false;
            }

            ExcelDataValues = officeAllocationFileData;
            //we update the display
            gSelectedFiles.Visibility = Visibility.Collapsed;
            gIntermediateData.Visibility = Visibility.Visible;
            gIntermediateData.ItemsSource = "";
            gIntermediateData.ItemsSource = new List<DataValue>() { new DataValue() {
                AgeGroup ="-",
                 FacilityName="",IndicatorId="File processed. Click Save to Database",IndicatorValue=0,ProgramArea="Timesheets",ReportMonth=singlefile.ReportMonth,ReportYear=singlefile.ReportYear,Sex=""

            } };// ExcelDataValues;
            gIntermediateData.Height = 500;
            return true;
        }

        private LocationDetail GetReportYearAndMonthFromFileNames(string fileName, int thisYear)
        {
            var fileNameParts = System.IO.Path.GetFileNameWithoutExtension(fileName).Split(new[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            if (fileNameParts.Length < 2)
                return null;

            var yearStr = fileNameParts[1];
            int reportYear;
            if (!int.TryParse(yearStr, out reportYear) || !Constants.acceptableYears.Contains(reportYear) || reportYear > thisYear)
            {
                MessageBox.Show("Error determining Report year or month. Please ensure the file ends with MMM YYYY");
                return null;
            }

            var monthStr = fileNameParts[0];
            var reportMonth = Constants.GetStandardMonthName(monthStr);
            if (string.IsNullOrWhiteSpace(reportMonth))
            {
                MessageBox.Show("Error determining Report year or month. Please ensure the file ends with MMM YYYY");
                return null;
            }
            var toReturn = new LocationDetail() { ReportYear = reportYear, ReportMonth = reportMonth };
            return toReturn;
        }

        private void clearSelected(object sender, RoutedEventArgs e)
        {
        }

        void saveToAltServer(object sender, RoutedEventArgs e)
        {
            SaveToServer(true);
        }

        private void saveToServer(object sender, RoutedEventArgs e)
        {
            SaveToServer(false);
        }

        void SaveToServer(bool saveToDevServer)
        {
            //we get valuesDataset
            if (ExcelDataValues == null || ExcelDataValues.Count == 0)
            {
                return;
            }

            var valuesDataset = ExcelDataValues.ToDataset();
            if (valuesDataset.Tables.Count == 0)
            {
                MessageBox.Show("Nothing to export");
                return;
            }
            var tempTableName = new RandomTableNameGenerator().Execute();
            valuesDataset.Tables[0].TableName = tempTableName;

            var currentProject = CurrentProjectName;
            var connBuilder = DbFactory.GetDefaultConnection(currentProject, saveToDevServer);
            if (connBuilder == null)
                return;

            var contextDb = new DbHelper(connBuilder);

            try
            {
                var dataImporter = new SaveTableToDbCommand()
                {
                    TargetDataset = valuesDataset,
                    Db = contextDb
                };

                dataImporter.Execute();

                //we start the merge
                var dataMerge = new TimesheetsDataMergeCommand()
                {
                    TempTableName = tempTableName,
                    DestinationTable = "TimesheetData",
                    Db = contextDb
                };

                // we save, 
                dataMerge.Execute();
            }
            catch
            {
                throw;
            }
            finally
            {
                tHelpfulTip.Content = "Select Files to Import";
                ResetAllGrids();
            }
        }
    }
}
