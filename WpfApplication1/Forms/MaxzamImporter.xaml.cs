using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
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
    /// Interaction logic for TimesheetImporter.xaml
    /// </summary>
    public partial class MaxzamImporter : Page
    {
        public MaxzamImporter()
        {
            InitializeComponent();
            SelectedFiles = new List<FinanceDetails>();
        }

        public ProjectName CurrentProjectName
        {
            get { return ProjectName.Maxzam; }
            internal set { }
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
                Multiselect = true,
                CheckPathExists = true,
                Filter = "CSV (*.csv)|*.csv",
                Title = "Please select the files to import"
            };
            var dialogResult = dialog.ShowDialog() ?? false;
            if (dialogResult)
            {
                foreach(var filename in dialog.FileNames)
                {
                    //we process the file
                    //we process for report year and month
                    SelectedFiles.Add(new FinanceDetails()
                    {
                        FileName = filename,
                        ReportYear = 0,
                        ReportMonth = ""
                    });
                }
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
            //ExcelDataValues.Clear();
        }

        double defaultGridSize = 500;
        private void reviewSelectedFiles(object sender, RoutedEventArgs e)
        {
            if (SelectedFiles.Count == 0)
                return;

            //fire up template reader and let it do the rest
            var selectedFiles = SelectedFiles;
            //var res = ReadDataFiles(SelectedFiles, CurrentProjectName);
            //if (res)
            //{
            //    SelectedFiles.Clear();
            //}
        }

        CodeRunner<List<DataValue>> _runner;

        bool ReadDataFiles(string fileName)
        {
            //static DataTable GetDataTableFromCsv(string path, bool isFirstRowHeader)
            //{

            //}

            //var officeAllocationFileData = new GetExcelAsDataTable2()
            //{
            //    reportYear = singlefile.ReportYear,
            //    reportMonth = singlefile.ReportMonth,
            //    locationDetail = null,
            //    Db = new DbHelper(DbFactory.GetDefaultConnection(CurrentProjectName)),
            //    fileName = singlefile.FileName,
            //    worksheetName = "Sheet1",
            //    progressDisplayHelper = new WaitDialog()
            //    {
            //        WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner
            //    },
            //    SelectedProject = ProjectName.General
            //}.Execute();

            //if (officeAllocationFileData == null)
            //{
            //    //we skip and alert the user of the error
            //    return false;
            //}

            //ExcelDataValues = officeAllocationFileData;

            //we update the display
            gSelectedFiles.Visibility = Visibility.Collapsed;
            gIntermediateData.Visibility = Visibility.Visible;
            gIntermediateData.ItemsSource = "";
            gIntermediateData.ItemsSource = new List<DataValue>() { new DataValue() {
                AgeGroup ="-",
                 FacilityName="",IndicatorId="File processed. Click Save to Database",
                IndicatorValue =0,
                ProgramArea ="Timesheets",
                //ReportMonth =singlefile.ReportMonth,
                //ReportYear =singlefile.ReportYear,
                Sex =""

            } };// ExcelDataValues;
            gIntermediateData.Height = 500;
            return true;
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
            var currentProject = CurrentProjectName;
            var connBuilder = DbFactory.GetDefaultConnection(currentProject, saveToDevServer);
            if (connBuilder == null)
                return;

            var contextDb = new DbHelper(connBuilder);

            var valuesDataset = new DataSet();
            var saveToDbHelper = new List<SaveTableToDbCommand>();

            var files = SelectedFiles.Select(t => t.FileName).ToList();
            var target_to_temp = new Dictionary<string, string>();
            foreach (var file in files)
            {
                var filenameonly = System.IO.Path.GetFileNameWithoutExtension(file);
                var ds_name = filenameonly.Substring(0, filenameonly.LastIndexOf('-')).Replace('-', '_').Replace("extract_","");
                var rndmname = new RandomTableNameGenerator().Execute();

                var csvLoader = new GetCsvData()
                {
                    fileName = file,Db= contextDb
                };

                var table = csvLoader.Execute();
                table.TableName = rndmname;
                valuesDataset.Tables.Add(table);

                target_to_temp[ds_name] = rndmname;
            }

            var stepx = 2;
            try
            {
                var dataImporter = new SaveTableToDbCommand()
                {
                    TargetDataset = valuesDataset,
                    Db = contextDb
                };

                dataImporter.Execute();

                //we start the merge
                var dataMerge = new MaxzamDataMergeCommand()
                {
                    destinationTempNames= target_to_temp,
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
        
        public class GetCsvData : IQueryHelper<DataTable>
        {
            public string fileName { get; set; }

            public DbHelper Db { get; set; }
            public IDisplayProgress progressDisplayHelper { get; set; }
            public Action<string> Alert { get; set; }
            
            public DataTable Execute()
            {
                //https://stackoverflow.com/questions/1050112/how-to-read-a-csv-file-into-a-net-datatable
                var toReturn = new DataTable();
                toReturn.Locale = CultureInfo.CurrentCulture;
               

                var pathOnly = System.IO.Path.GetDirectoryName(fileName);

                var fileonly = System.IO.Path.GetFileName(fileName);

                var m = new OleDbConnectionStringBuilder(
                    "Provider=Microsoft.Jet.OLEDB.4.0; Data Source="+ fileName + ";Extended Properties=\"Text;HDR=Yes;IMEX=1\"");

                //Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1""

                var sql = @"SELECT * FROM [" + fileonly + "]";
                using (var connection = new OleDbConnection(
                          @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathOnly +
                          ";Extended Properties=\"Text;HDR=Yes\""))
                using (var command = new OleDbCommand(sql, connection))
                using (var adapter = new OleDbDataAdapter(command))
                {
                    adapter.Fill(toReturn);
                }
                return toReturn;
            }

            private DataTable ImportData()
            {
                return new DataTable();
            }
        }
    }
}
