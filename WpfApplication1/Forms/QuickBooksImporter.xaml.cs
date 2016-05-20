using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
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
    /// Interaction logic for QuickBooksImporter.xaml
    /// </summary>
    public partial class QuickBooksImporter : Page
    {
        public QuickBooksImporter()
        {
            InitializeComponent();
            SelectedFiles = new List<FileDetails>();
        }

        ProjectName CurrentProjectName = ProjectName.General;
        public List<FileDetails> SelectedFiles { get; private set; }

        private void selectFile(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog()
            {
                CheckFileExists = true,
                Multiselect = true,
                CheckPathExists = true,
                Filter = "Excel (*.xlsx)|*.xlsx|Excel (*.xls)|*.xls",
                Title = "Please select the files to import"
            };
            var dialogResult = dialog.ShowDialog() ?? false;
            if (dialogResult)
            {
                if (dialog.FileNames.Length != 2)
                {
                    MessageBox.Show("Please only select the 2 Quick Book files for a given month",
                        "Incorrect nummber of files selected", MessageBoxButton.OK);
                        return;
                }

                SelectedFiles.AddRange(
                    (from file in dialog.FileNames
                     select new FileDetails() { FileName = file }).ToList()
                    );
            }

            //we prompt for the date
            foreach(var fileDetail in SelectedFiles)
            {
                var filename = fileDetail.FileName.Split(' ');

            }

            //we refresh the grid
            gIntermediateData.Visibility = Visibility.Collapsed;

            gSelectedFiles.ItemsSource = "";
            gSelectedFiles.ItemsSource = SelectedFiles;
            gSelectedFiles.Visibility = Visibility.Visible;
            //refreshDataGrid(true);

            //update the tip
            tHelpfulTip.Content = "Click Select Files to add more files or select next option";
        }

        void EnableSaveButtons(bool viewState) { }

        private void ShowGridDisplayPort(List<DataValue> table1, List<DataValue> table2)
        {
            //refreshDataGrid(false);
            gIntermediateData.ItemsSource = "";
            gIntermediateData.ItemsSource = table1;
        }

        private void ResetAllGrids()
        {
            gSelectedFiles.Height = 0;
            gIntermediateData.Height = 0;
            gSelectedFiles.ItemsSource = "";
            gIntermediateData.ItemsSource = "";
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
            ReadDataFiles(SelectedFiles, CurrentProjectName);
        }

        CodeRunner<List<DataValue>> _runner;
        public volatile List<MatchedDataValue> ExcelDataValues = null;

        public class FilesSelectedInfo
        {
            public string OfficeAllocationFile;
            public string TotalCostsFile;

            public LocationDetail LocationDetails;
        }

        public FilesSelectedInfo AssignSelectedFiles(List<FileDetails> files)
        {
            var thisYear = DateTime.Now.Year;

            var officeAllocationFile = string.Empty;
            var totalCostsFile = string.Empty;
            var reportYear = -1;
            var reportMonth = string.Empty;

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
                var reportYearAndMonth = GetReportYearAndMonthFromFileNames(filename, thisYear);
                if (reportYearAndMonth == null)
                {
                    return null;
                }
                if (i > 0 && reportYearAndMonth.ReportYear != reportYear && reportYearAndMonth.ReportMonth != reportMonth)
                {
                    MessageBox.Show("The two files have inconsistent dates");
                    return null;
                }
                reportYear = reportYearAndMonth.ReportYear;
                reportMonth = reportYearAndMonth.ReportMonth;
            }
            return new FilesSelectedInfo()
            {
                OfficeAllocationFile = officeAllocationFile,
                LocationDetails = new LocationDetail() { ReportMonth = reportMonth, ReportYear = reportYear },
                TotalCostsFile = totalCostsFile
            };
        }

        void ReadDataFiles(List<FileDetails> files, ProjectName projectName)
        {
            if (files.Count != 2)
                throw new ArgumentOutOfRangeException("Expected two files passed in");

            var processedFilesInfo = AssignSelectedFiles(files);
            if (processedFilesInfo == null)
                return;

            processedFilesInfo.LocationDetails.FacilityName = "5040HQ5";
            if (ExcelDataValues == null) { ExcelDataValues = new List<MatchedDataValue>(); } else { ExcelDataValues.Clear(); }
            var officeAllocationFileData = new GetFinanceDataFromExcel()
            {
                locationDetail = processedFilesInfo.LocationDetails,
                fileName = processedFilesInfo.OfficeAllocationFile,
                worksheetName =Constants.INCOUNTRY_ION_EXPENSES,
                progressDisplayHelper = new WaitDialog()
                {
                    WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner
                },
                SelectedProject = ProjectName.General
            }.Execute();

            var totalCostsData = new GetFinanceDataFromExcel()
            {
                locationDetail = processedFilesInfo.LocationDetails,
                fileName = processedFilesInfo.TotalCostsFile,
                worksheetName = string.Empty,
                progressDisplayHelper = new WaitDialog()
                {
                    WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner
                },
                SelectedProject = ProjectName.General
            }.Execute();

            if (officeAllocationFileData == null || totalCostsData == null)
            {
                //we skip and alert the user of the error
                return;
            }

            var dict = new Dictionary<string, TwoDataValuePair>();
            foreach(var totalCost in totalCostsData)
            {
                dict[totalCost.IndicatorId+ totalCost.AgeGroup] = new TwoDataValuePair()
                {
                    TotalCostDataValue = totalCost
                };
            }

            foreach (var officeAlloc in officeAllocationFileData)
            {
                TwoDataValuePair twoValuePair;
                var matching = dict.TryGetValue(officeAlloc.IndicatorId+ officeAlloc.AgeGroup, out twoValuePair);
                if (twoValuePair == null)
                {
                    twoValuePair = new TwoDataValuePair();
                    dict[officeAlloc.IndicatorId + officeAlloc.AgeGroup] = twoValuePair;
                }
                twoValuePair.OfficeAllocationDataValue = officeAlloc;
            }

            var table1 = (from matchedDvs in dict.Values
                          //let x = matchedDvs.AsMatchedDataValue()
                          //where x.OfficeAllocation != 0 && x.TotalCost != 0
                          //select x
                          select matchedDvs.AsMatchedDataValue()
                          ).ToList();

            //we update the display
            gSelectedFiles.Visibility = Visibility.Collapsed;
            gIntermediateData.Visibility = Visibility.Visible;
            gIntermediateData.ItemsSource = "";
            gIntermediateData.ItemsSource = table1;
            gIntermediateData.Height = 500;
            //ShowGridDisplayPort(table1, table2);
            return;
        }

        private LocationDetail GetReportYearAndMonthFromFileNames(string fileName, int thisYear)
        {
            var fileNameParts = System.IO.Path.GetFileNameWithoutExtension(fileName).Split(new[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            if (fileNameParts.Length < 2)
                return null;

            var yearStr = fileNameParts[fileNameParts.Length - 1];
            int reportYear;
            if (!int.TryParse(yearStr, out reportYear) || !Constants.acceptableYears.Contains(reportYear) || reportYear > thisYear)
            {
                MessageBox.Show("Error determining Report year or month. Please ensure the file ends with MMM YYYY");
                return null;
            }

            var monthStr = fileNameParts[fileNameParts.Length - 2];
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

            var valuesDataset = ExcelDataValues .ToDataset();
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
                var dataMerge = new DataMergeCommand()
                {
                    TempTableName = tempTableName,
                    DestinationTable = "FacilityData",
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
