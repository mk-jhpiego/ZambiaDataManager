using System;
using System.Collections.Generic;
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
using ZambiaDataManager.Popups;
using ZambiaDataManager.Storage;

namespace ZambiaDataManager
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void showMap(object sender, RoutedEventArgs e)
        {
            //we pick he folder
            var baseFolderName = "C:\\Data Manager Files";
            const string added = "added";
            const string notAdded = "notadded";
            var folders = new List<string>() {baseFolderName,
                System.IO.Path.Combine(baseFolderName, added),
                System.IO.Path.Combine(baseFolderName, notAdded)
            };

            foreach (var folder in folders)
            {
                if (!Directory.Exists(baseFolderName))
                {
                    Directory.CreateDirectory(baseFolderName);
                }
            }

            //and start monitoring files or run everyonce in a while
            while (true)
            {
                //we select the file
                var unprocessedFiles = Directory.GetFiles(
                    baseFolderName, "*.xlsz");

                foreach(var file in unprocessedFiles)
                {
                    //

                }
            }

            //for files that we've processed, we move them to a separate folder

            //
        }

        private void showPepfarReport(object sender, RoutedEventArgs e)
        {

        }

        private void showProgramIndicatorsReport(object sender, RoutedEventArgs e)
        {

        }

        private void showAllIndicatorsReport(object sender, RoutedEventArgs e)
        {

        }

        private void viewAllSitesReporting(object sender, RoutedEventArgs e)
        {

        }

        private void addVmmcMonthly(object sender, RoutedEventArgs e)
        {
            var targetForm = new Forms.pageAddMERLData()
            {
                CurrentProjectName = ProjectName.IHP_VMMC
            };
            stackMain.Content = targetForm;
        }
        private void addDodMonthly(object sender, RoutedEventArgs e)
        {
            var targetForm = new Forms.pageAddMERLData()
            {
                CurrentProjectName = ProjectName.DOD
            };
            stackMain.Content = targetForm;
        }

        private void addVmmcCampaignDailyData(object sender, RoutedEventArgs e)
        {
            var targetForm = new Forms.pageAddMERLData()
            {
                CurrentProjectName = ProjectName.IHP_VMMC
            };
            stackMain.Content = targetForm;
        }

        private void reviewUploadedData(object sender, RoutedEventArgs e)
        {
            var targetForm = new Forms.pageReviewData();
            stackMain.Content = targetForm;
        }

        private void addVmmcTechnicalReport(object sender, RoutedEventArgs e)
        {

        }

        private void manageAccount(object sender, RoutedEventArgs e)
        {

        }

        private void updateLocalRepo(object sender, RoutedEventArgs e)
        {

        }

        System.Data.DataTable refreshDataHandler(
            string procName,            
            List<KeyValuePair<string, object>> parameters
            )
        {
            var dbconnection = DbFactory.GetDefaultConnection(
                PageController.Instance.DefaultProjectName);
            var db = new DbHelper(dbconnection);
            var table = db.GetTable(procName, true, parameters);
            return table;
        }

        private void bViewPepfarReport_Click(object sender, RoutedEventArgs e)
        {
            var procName = "proc_getcoag_report";
            var parameters = new List<KeyValuePair<string, object>>() {
                    new KeyValuePair<string, object>( "@yearMonth","2017 Mar") };
            var table = refreshDataHandler(procName, parameters);
            var targetForm = new Forms.gridDisplay()
            {
                procedureName = procName,
                refreshHandler = refreshDataHandler
            };
            targetForm.refreshDataGrid(table);
            stackMain.Content = targetForm;
        }

        private void showPRSreport(object sender, RoutedEventArgs e)
        {
            var procName = "proc_get_summary_pivot";
            var parameters = new List<KeyValuePair<string, object>>() {
                    new KeyValuePair<string, object>( "@yearMonth","2017 Mar") };
            var table = refreshDataHandler(procName, parameters);
            var targetForm = new Forms.gridDisplay()
            {
                procedureName = procName,
                refreshHandler = refreshDataHandler
            };
            targetForm.refreshDataGrid(table);
            stackMain.Content = targetForm;
        }

        Forms.waitWindow _waitWindow = null;
        void setIsReady(bool status, string msg = "")
        {
            stackUserMenu.IsEnabled = status;
            if (_waitWindow == null)
            {
                _waitWindow= new Forms.waitWindow() { showWait = status, displayMsg = msg };
                stackMain.Content = _waitWindow;
            }
            _waitWindow.showWait = status;
            _waitWindow.displayMsg = msg;
            _waitWindow.InvalidateArrange();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Application.Current.DispatcherUnhandledException += Current_DispatcherUnhandledException;

            //show the wait message
            
            //we prompt for the default database
            var defaultProject = mainapp.Instance.Worker.getDefaultProject();

            if (defaultProject == ProjectName.None)
            {
                var dialog = new ProjectedSelector();
                if (dialog.ShowDialog() != null && dialog.SelectedProjectName != ProjectName.None)
                {
                    defaultProject = dialog.SelectedProjectName;
                    if (dialog.RememberSelection)
                    {
                        //we save the selected project
                        mainapp.Instance.Worker.setDefaultProject(defaultProject);
                    }
                }
                else
                {
                    //means they canceled
                    //we disable everything
                    Application.Current.Shutdown();
                }
            }

            PageController.Instance.DefaultProjectName = defaultProject;
            Title = "Jhpiego Zambia Data Manager v2.1: Default project selected is " + defaultProject.ToString();
            var user = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            tLoggedInUser.Text = user ?? "Not Logged In";

            //we show the wait window
            setIsReady(false);
            
            _syncClass = new syncClass();
            var scheduler = TaskScheduler
                .FromCurrentSynchronizationContext();
            var task = Task.Run(()=>setUpDatabase()).
                ContinueWith(t=> setIsReady(_syncClass.successful, _syncClass.message), scheduler);
        }

        volatile syncClass _syncClass;

        class syncClass
        {
            public bool successful { get; set; }
            public string message { get; set; }
        }

        void setUpDatabase()
        {
            //we load db data
            var shutDown = false;
            //we check the connection
            if (!isConnected())
            {
                var dbconnection = DbFactory.GetDefaultConnection(
                    PageController.Instance.DefaultProjectName);
                _syncClass.message = string.Format(
                    "Could not connect to the server {0}\\{1}",
                    dbconnection.ServerName, dbconnection.DatabaseName);
                _syncClass.successful = false;
                return;
            }

            try
            {
                loadDbData();
            }
            catch (SqlException sqlex)
            {
                _syncClass.message = "Could not connect to the database. The application will shut down";
                //MessageBox.Show("Could not connect to the database. The application will shut down");
                shutDown = true;
            }
            catch (Exception ex)
            {
                _syncClass.message = "Could not start the application. Error has been logged. The application will shut down";
                //MessageBox.Show("Could not start the application. Error has been logged. The application will shut down");
                shutDown = true;
            }
            finally
            {
            }

            _syncClass.successful = !shutDown;
            return;
        }

        private void Current_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            MessageBox.Show("An issue is preventing startup. Please retry. The app will shutdown. If the problem persists, please call for support");
            e.Handled = true;
            Application.Current.Shutdown();
        }

        bool isConnected()
        {
            var toReturn = false;
            var dbconnection = DbFactory.GetDefaultConnection(
    PageController.Instance.DefaultProjectName);
            var db = new DbHelper(dbconnection);
            try
            {
                var test = db.GetScalar("select 1", 30);
                toReturn = true;
            }
            catch (SqlException e)
            {
                toReturn = false;
            }
            catch (Exception e)
            {
                toReturn = false;
            }
            return toReturn;
        }

        bool loadDbData()
        {
            var toReturn = false;
            //what db data
            var dbconnection = DbFactory.GetDefaultConnection(
                PageController.Instance.DefaultProjectName);         
            var db = new DbHelper(dbconnection);
            var provider = new AgegroupsProvider()
            {
                DB = db
            };
            //
            var alternateAgeGroups = provider.getAlternateAgeGroups();
            var cleanAges = new Dictionary<string, string>();
            foreach(var age in alternateAgeGroups)
            {
                //there might be some overwriting but its fine
                cleanAges[age.Key.toCleanAge()] = age.Value;
            }

            PageController.Instance.AlternateAgegroups = cleanAges;
            return toReturn;
        }

        private void addQuickBooksData(object sender, RoutedEventArgs e)
        {
            var targetForm = new Forms.QuickBooksImporter()
            {
                CurrentProjectName = ProjectName.General
            };
            stackMain.Content = targetForm;
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            Application.Current.Shutdown();
        }
    }
}
