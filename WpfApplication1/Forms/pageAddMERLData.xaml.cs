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
    /// Interaction logic for pageAddMERLData.xaml
    /// </summary>
    public partial class pageAddMERLData : Page
    {
        public pageAddMERLData()
        {
            InitializeComponent();
            SelectedFiles = new List<FileDetails>();
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
                SelectedFiles.AddRange(
                    (from file in dialog.FileNames
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

        private void ResetAllGrids()
        {
            _currentProjectName = ProjectName.None;
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
            ReadDataFiles(SelectedFiles, ProjectName.IHP_VMMC);
        }

        CodeRunner<List<DataValue>> _runner;
        public volatile List<DataValue> ExcelDataValues = null;
        ProjectName _currentProjectName = ProjectName.None;

        void ReadDataFiles(List<FileDetails> files, ProjectName projectName)
        {
            if (ExcelDataValues == null){ExcelDataValues = new List<DataValue>();}else { ExcelDataValues.Clear(); }
            _currentProjectName = projectName;

            foreach (var file in files)
            {
                var worker = new GetValuesFromReport() { fileName = file.FileName, SelectedProject = projectName };
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


            var fileName = string.Empty;
            if (!string.IsNullOrWhiteSpace(fileName) && File.Exists(fileName))
            {
                _runner = new CodeRunner<List<DataValue>>()
                {
                    ShowSplash = true,
                    CodeToExcute = new GetValuesFromReport() { fileName = fileName, SelectedProject = projectName },
                    AsyncCallBack = (q) =>
                    {
                        if (q == null)
                            return;

                        EnableSaveButtons(true);
                        ExcelDataValues = q;
                        ////valuesDataset = q.ToDataset();
                        //if (dataGridView1.InvokeRequired)
                        //{
                        //    dataGridView1.Invoke(
                        //        new refreshDisplay((s) => {
                        //            ShowGridDisplayPort(s.Tables[0]);
                        //        }),
                        //        valuesDataset);
                        //    return;
                        //}
                        ShowGridDisplayPort(ExcelDataValues);
                    }
                };
                _runner.Execute();
            }
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


            var currentProject = _currentProjectName;
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
