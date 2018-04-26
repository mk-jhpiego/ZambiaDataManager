using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
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
    /// Interaction logic for pageAddMERLData.xaml
    /// </summary>
    public partial class pageAddMERLData : Page
    {
        public pageAddMERLData()
        {
            InitializeComponent();
            SelectedFiles = new List<FileDetails>();
        }

        public ProjectName CurrentProjectName
        {
            get;
            internal set;
        }

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
                var currentFiles = (from file in SelectedFiles select file.FileName).ToList();
                var dialogFiles = dialog.FileNames.ToList().Except(currentFiles);
                SelectedFiles.AddRange(
                    (from file in dialogFiles
                     select new FileDetails() { FileName = file }).ToList()
                    );
            }
            //we refresh the grid
            refreshDataGrid(true);

            //update the tip
            tHelpfulTip.Content = "Click Select Files to add more files or select next option";
        }

        void refreshDataGrid(bool showSelectedFiles)
        {
            gSelectedFiles.Height = 0;
            gIntermediateData.Height = 0;
            if (showSelectedFiles)
            {
                gSelectedFiles.Height = defaultGridSize;
                gSelectedFiles.ItemsSource = "";
                gSelectedFiles.ItemsSource = SelectedFiles;
            }
            else
            {
                gIntermediateData.Height = defaultGridSize;

            }
        }

        void EnableSaveButtons(bool viewState){}

        private void ShowGridDisplayPort(List<DataValue> dataSource = null)
        {
            refreshDataGrid(false);
            gIntermediateData.ItemsSource = "";
            gIntermediateData.ItemsSource = dataSource;
        }
        public BaseMergeCommand mergeHelper { get; set; }
        public DataTable webData = null;
        public void ShowGridDisplayPort(DataTable dataSource = null)
        {
            refreshDataGrid(false);
            gIntermediateData.ItemsSource = "";
            gIntermediateData.Columns.Clear();
            gIntermediateData.AutoGenerateColumns = true;
            gIntermediateData.IsReadOnly = true;
            gIntermediateData.HorizontalScrollBarVisibility = ScrollBarVisibility.Visible;
           // gIntermediateData.sc
            bSelectFile.IsEnabled = false;
            bReviewFiles.IsEnabled = false;
            webData = dataSource;
            gIntermediateData.ItemsSource = webData.DefaultView;
        }

        private void ResetAllGrids()
        {
            gSelectedFiles.Height = 0;
            gIntermediateData.Height = 0;
            gSelectedFiles.ItemsSource = "";
            gIntermediateData.ItemsSource = "";
            SelectedFiles.Clear();
            webData = null;
            if (ExcelDataValues != null) ExcelDataValues.Clear();
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
        public volatile List<DataValue> ExcelDataValues = null;
        //ProjectName _currentProjectName = ProjectName.None;

        void ReadDataFiles(List<FileDetails> files, ProjectName projectName)
        {
            if (ExcelDataValues == null){ExcelDataValues = new List<DataValue>();}else { ExcelDataValues.Clear(); }
            var connBuilder = DbFactory.GetDefaultConnection(CurrentProjectName);
            if (connBuilder == null)
                return;

            foreach (var file in files)
            {
                IQueryHelper<List<DataValue>> worker = null;
                if (projectName == ProjectName.DOD)
                {
                    worker = new GetDodDataFromExcel()
                    {
                        fileName = file.FileName,
                        SelectedProject = projectName,
                        ageGroupsProvider = new AgegroupsProvider() { DB = new DbHelper(connBuilder) }
                        //locationDetail = processedFilesInfo.LocationDetails,
                        //worksheetName = Constants.INCOUNTRY_ION_EXPENSES
                    };
                }
                else
                {
                    worker = new GetValuesFromReport() { fileName = file.FileName, SelectedProject = projectName };
                }
                worker.progressDisplayHelper = new WaitDialog() { WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner };

                var intermediateResults = worker.Execute();
                if (intermediateResults == null)
                {
                    //we skip and alert the user of the error
                    continue;
                }

                ExcelDataValues.AddRange(intermediateResults);
            }

            //we update the display
            ShowGridDisplayPort(ExcelDataValues);
            return;            
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
            if ((ExcelDataValues == null || ExcelDataValues.Count == 0) && webData == null)
            {
                return;
            }

            DataSet valuesDataset = null;
            if (webData != null)
            {
                valuesDataset = new DataSet();
                //webData.TableName = "DataValue";
                valuesDataset.Tables.Add(webData);
            }
            else
            {
                valuesDataset = ExcelDataValues.ToDataset();
            }
            if (valuesDataset.Tables.Count == 0)
            {
                MessageBox.Show("Nothing to export");
                return;
            }
            var tempTableName = new RandomTableNameGenerator().Execute();
            valuesDataset.Tables[0].TableName = tempTableName;

            var connBuilder = DbFactory.GetDefaultConnection(CurrentProjectName, saveToDevServer);
            if (connBuilder == null)
                return;

            var contextDb = new DbHelper(connBuilder);
            try
            {
                var dataImporter = new SaveTableToDbCommand()
                {
                    TargetDataset = valuesDataset,
                    Db = contextDb,
                    IsWebData = (webData != null)
                };

                dataImporter.Execute();
            }
            catch(SqlException sqlex)
            {
                throw;
            }
            catch
            {
                throw;
            }

            try
            {
                //we start the merge
                var dataMerge = mergeHelper;
                dataMerge.TempTableName = tempTableName;
                dataMerge.DestinationTable = "FacilityData";
                dataMerge.Db = contextDb;
                dataMerge.IsWebData = (webData != null);

                // we save, 
                dataMerge.Execute();
            }
            catch(Exception ex)
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
