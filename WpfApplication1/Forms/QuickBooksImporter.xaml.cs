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

        //void refreshDataGrid(bool showSelectedFiles)
        //{
        //    gSelectedFiles.Height = 0;
        //    gIntermediateData.Height = 0;
        //    if (showSelectedFiles)
        //    {
        //        gSelectedFiles.Height = defaultGridSize;
        //        gSelectedFiles.ItemsSource = "";
        //        gSelectedFiles.ItemsSource = SelectedFiles;
        //    }
        //    else
        //    {
        //        gIntermediateData.Height = defaultGridSize;

        //    }
        //}

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
        public volatile List<DataValue> ExcelDataValues = null;

        void ReadDataFiles(List<FileDetails> files, ProjectName projectName)
        {
            //Spending Data Table containing differences
            //Sheet1 in TOtal spending for month file
            //'Expenses for expns for the mnth' for the mnth in Office allocation

            //Incountry IONs expenses for total expenses
            //QB Office Allocation for TPI for office allocations by project

            if (ExcelDataValues == null) { ExcelDataValues = new List<DataValue>(); } else { ExcelDataValues.Clear(); }
            var table1 = new GetFinanceDataFromExcel()
            {
                fileName = files[0].FileName,
                progressDisplayHelper = new WaitDialog()
                {
                    WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner
                }
                ,
                SelectedProject = ProjectName.General
            }.Execute();

            //var table2 = new GetFinanceDataFromExcel()
            //{
            //    fileName = files[1].FileName,
            //    progressDisplayHelper = new WaitDialog()
            //    {
            //        WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner
            //    },
            //    SelectedProject = ProjectName.General
            //}.Execute();
            var table2 = new { };
            if (table1 == null || table2 == null)
            {
                //we skip and alert the user of the error
                return;
            }

            //we update the display
            gSelectedFiles.Visibility = Visibility.Collapsed;
            gIntermediateData.Visibility = Visibility.Visible;
            gIntermediateData.ItemsSource = "";
            gIntermediateData.ItemsSource = table1;
            gIntermediateData.Height = 500;
            //ShowGridDisplayPort(table1, table2);
            return;
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
