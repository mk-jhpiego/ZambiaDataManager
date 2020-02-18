using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using ZambiaDataManager.modules;
using ZambiaDataManager.modules.mcsp_lng;
using ZambiaDataManager.modules.vmmc;
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
            Title = "Jhpiego Zambia Data Manager 2020.02.18: Default project selected is " + defaultProject.ToString();
            var user = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            tLoggedInUser.Text = user ?? "Not Logged In";

            //we show the wait window
            setIsReady(false);

            _syncClass = new syncClass();
            var scheduler = TaskScheduler
                .FromCurrentSynchronizationContext();
            var task = Task.Run(() => setUpDatabase()).
                ContinueWith(t => setIsReady(_syncClass.successful, _syncClass.message), scheduler);
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

        private void showMap(object sender, RoutedEventArgs e)
        {
            return;
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
            return;

            //var targetForm = new Forms.pageAddMERLData()
            //{
            //    CurrentProjectName = ProjectName.IHP_VMMC
            //};
            //stackMain.Content = targetForm;
        }
        private void addDodMonthly(object sender, RoutedEventArgs e)
        {
            var targetForm = new Forms.pageAddMERLData()
            {
                CurrentProjectName = ProjectName.DOD
            };
            stackMain.Content = targetForm;
        }

        //private void addVmmcCampaignDailyData(object sender, RoutedEventArgs e)
        //{
        //    var targetForm = new Forms.pageAddMERLData()
        //    {
        //        CurrentProjectName = ProjectName.IHP_VMMC
        //    };
        //    stackMain.Content = targetForm;
        //}

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
                    new KeyValuePair<string, object>( "@yearMonth","2017 Apr") };
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
            var filter = new Forms.pageYearMonthFilter()
            {
                FilterCallBack = (int nothing, string yearMonth) => {
                    loadPrsReport(yearMonth);
                    //var latestData = await downloadLatest(year, monthTxt);
                    //if (latestData != null)
                    //{
                    //    getVmmcWebData2(latestData);
                    //}
                }
            };
            stackMain.Content = filter;

            //loadPrsReport();
        }

        private void loadPrsReport(string yearMonth)
        {
            var procName = "proc_get_summary_pivot";
            var parameters = new List<KeyValuePair<string, object>>() {
                //yearMonth
                //"2017 Apr"
                    new KeyValuePair<string, object>( "@yearMonth",yearMonth) };
            var table = refreshDataHandler(procName, parameters);
            var targetForm = new Forms.gridDisplay()
            {
                procedureName = procName,
                refreshHandler = refreshDataHandler
            };
            targetForm.refreshDataGrid(table);
            stackMain.Content = targetForm;
        }

        private void bViewJadeQtrlyReport_Click(object sender, RoutedEventArgs e)
        {
            var procName = "proc_getjade_qtrly_2017q1";
            var parameters = new List<KeyValuePair<string, object>>() {
                    //new KeyValuePair<string, object>( "@yearMonth","2017 Mar")
            };
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
            if(ProjectName.Maxzam == PageController.Instance.DefaultProjectName)
            {
                return toReturn;
            }
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

        //async Task<string> downloadLatest(int year, string monthTxt)
        //{
        //    //todo: get file for a partcular year and month
        //    //https://firebasestorage.googleapis.com/v0/b/daily-reporting-4ac35.appspot.com/o/mcsp%2Flng%2Freceiving-iud%2Flatest.json?alt=media&token=711cff83-1762-4fef-94a2-69fc5ee7d99a
        //    //var baseUrl = "https://firebasestorage.googleapis.com/v0/b/daily-reporting-4ac35.appspot.com/o/monthlyreports%2F{0}?alt=media&token=a201912b-0ff4-4143-a00d-8a0b78791e82";
        //    var baseUrl = "https://firebasestorage.googleapis.com/v0/b/daily-reporting-4ac35.appspot.com/o/monthlyreports%2F{0}%2F{1}%2F{2}?alt=media&token=a201912b-0ff4-4143-a00d-8a0b78791e82";
        //    var latestfilelink = "latest.json";
        //    var currentFileLinkUrl = string.Format(baseUrl, year, monthTxt, latestfilelink);
        //    try
        //    {
        //        var res1 = await getWebData(currentFileLinkUrl);
        //        var latestfilename = Encoding.UTF8.GetString(res1);
        //        var dataFileUrl = string.Format(baseUrl, year, monthTxt, latestfilename);
        //        var reportData = await getWebData(dataFileUrl);
        //        //var reportText = Encoding.UTF8.GetString(reportData);
        //        File.WriteAllBytes(latestfilename, reportData);
        //        //Enable review and upload button
        //        return latestfilename;
        //    }
        //    catch (Exception x)
        //    {
        //        //labelStatus.Content = "Please specify a correct year";
        //        return null;
        //    }
        //}
        async Task<string> downloadLatest(string baseUrl, string resourcePath, string latestfilelink)
        {
            //var latestfilelink = "ppxlatest.json";
            var currentFileLinkUrl = string.Format(baseUrl, resourcePath, latestfilelink);
            try
            {
                var res1 = await getWebData(currentFileLinkUrl);
                var latestfilename = Encoding.UTF8.GetString(res1);
                var dataFileUrl = string.Format(baseUrl, resourcePath, latestfilename);
                var reportData = await getWebData(dataFileUrl);
                //var reportText = Encoding.UTF8.GetString(reportData);
                File.WriteAllBytes(latestfilename, reportData);
                //Enable review and upload button
                return latestfilename;
            }
            catch (Exception x)
            {
                //labelStatus.Content = "Please specify a correct year";
                return null;
            }
        }

        async Task<string> downloadLatest( string baseUrl, string resourcePath)
        {
            return await downloadLatest(baseUrl, resourcePath, "latest.json");

            //var latestfilelink = "latest.json";
            //var currentFileLinkUrl = string.Format(baseUrl, resourcePath, latestfilelink);
            //try
            //{
            //    var res1 = await getWebData(currentFileLinkUrl);
            //    var latestfilename = Encoding.UTF8.GetString(res1);
            //    var dataFileUrl = string.Format(baseUrl, resourcePath, latestfilename);
            //    var reportData = await getWebData(dataFileUrl);
            //    //var reportText = Encoding.UTF8.GetString(reportData);
            //    File.WriteAllBytes(latestfilename, reportData);
            //    //Enable review and upload button
            //    return latestfilename;
            //}
            //catch (Exception x)
            //{
            //    //labelStatus.Content = "Please specify a correct year";
            //    return null;
            //}
        }

        async Task<byte[]> getWebData(string url)
        {
            var http = new System.Net.Http.HttpClient();
            //var res1 = await http.GetByteArrayAsync(url);
            //var text = Encoding.UTF8.GetString(res1);
            return await http.GetByteArrayAsync(url);
        }

        public class AuthorisationResult
        {
            public AuthorisationRequest request { get; set; }
            public bool isAuthorised { get; set; }
        }

        public class AuthorisationRequest
        {
            public int resourceId { get; set; }
            public string userName { get; set; }
        }

        AuthorisationResult authoriseFor(AuthorisationRequest token)
        {
            return new AuthorisationResult()
            {
                isAuthorised = (token == null && (token.userName.ToLowerInvariant() != "global\\ambewe" || token.userName.ToLowerInvariant() != "global\\mkabila" || token.userName.ToLowerInvariant() != "global\\mnyambe")) ? false : true,
                request = token
            };
        }

        private async void getVmmcWebData2<T>(string fileName, BaseMergeCommand dataMerge) where T: class, ISerialisableWebDataset,new()
        {
            var user = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            var authorisationResult = authoriseFor(new AuthorisationRequest() { resourceId = Constants.ImportVmmcWeb, userName = user });
            if(authorisationResult==null|| !authorisationResult.isAuthorised)
            {
                MessageBox.Show("Access to this functionality is restricted", "Restricted access", MessageBoxButton.OK);
                return;
            }

            var jsonIndicatorData = File.ReadAllText(fileName);
            var indicators =
            Newtonsoft.Json.JsonConvert.DeserializeObject<List<T>>(jsonIndicatorData);

            var table = (new T()).getTable();

            indicators.ForEach(t => {
                var row = table.NewRow();
                table.Rows.Add(t.toRow(row));
            });

            var targetForm = new Forms.pageAddMERLData() {
                CurrentProjectName = dataMerge.projectName,
                mergeHelper= dataMerge
            };
            targetForm.ShowGridDisplayPort(table);

            stackMain.Content = targetForm;
        }

        private async void getVmmcWebData(object sender, RoutedEventArgs e)
        {
            var filter = new Forms.pageYearMonthFilter()
            {
                FilterCallBack = async (int year, string monthTxt) => {
                    //gs://daily-reporting-4ac35.appspot.com/monthlyreports/2019/10/monthly-20191118-183737.json
                    var baseUrl = "https://firebasestorage.googleapis.com/v0/b/daily-reporting-4ac35.appspot.com/o/monthlyreports%2F{0}%2F{1}?alt=media&token=a201912b-0ff4-4143-a00d-8a0b78791e82";
                    var webpath = String.Join("%2F", year, monthTxt);
                    var latestData = await downloadLatest(baseUrl,webpath);
                    if (latestData != null)
                    {
                        getVmmcWebData2<vmmcIndicatorRow>(latestData, 
                            dataMerge: new VmmcDataMergeCommand() {
                                facilityDataName = "FacilityData",
                                datasetName = Constants.ProjectTerms.VMMC, projectName = ProjectName.IHP_VMMC });
                    }
                    else
                    {
                    }
                }
            };
            stackMain.Content = filter;
        }

        private async void getPpxWebData(object sender, RoutedEventArgs e)
        {
            var filter = new Forms.pageYearMonthFilter()
            {
                FilterCallBack = async (int year, string monthTxt) => {
                    var baseUrl = "https://firebasestorage.googleapis.com/v0/b/daily-reporting-4ac35.appspot.com/o/monthlyreports%2F{0}%2F{1}?alt=media&token=a201912b-0ff4-4143-a00d-8a0b78791e82";
                    var webpath = String.Join("%2F", year, monthTxt);
                    var latestData = await downloadLatest(baseUrl, webpath, "ppxlatest.json");
                    if (latestData != null)
                    {
                        getVmmcWebData2<vmmcIndicatorRow>(latestData,
                            dataMerge: new VmmcDataMergeCommand()
                            {
                                facilityDataName="FacilityDataPpx",
                                datasetName = Constants.ProjectTerms.VMMC,
                                projectName = ProjectName.IHP_VMMC
                            });
                    }
                    else
                    {
                    }
                }
            };
            stackMain.Content = filter;
        }
        private async void getReceivingLngWebData(object sender, RoutedEventArgs e)
        {
            var baseUrl =
   "https://firebasestorage.googleapis.com/v0/b/daily-reporting-4ac35.appspot.com/o/mcsp%2Flng%2F{0}%2F{1}?alt=media&token=711cff83-1762-4fef-94a2-69fc5ee7d99a";
            var datasetname = Constants.ProjectTerms.RECEIVING_IUD;
            getLngData(baseUrl, datasetname);
        }

        private async void getDiscontinueLngWebData(object sender, RoutedEventArgs e)
        {
            var baseUrl =
                    "https://firebasestorage.googleapis.com/v0/b/daily-reporting-4ac35.appspot.com/o/mcsp%2Flng%2F{0}%2F{1}?alt=media&token=711cff83-1762-4fef-94a2-69fc5ee7d99a";
            var datasetname = Constants.ProjectTerms.DISCONTINUE_IUD;
            getLngData(baseUrl, datasetname);
        }

        private async void getLngData(string baseUrl, string datasetname)
        {
            var filter = new Forms.pageYearMonthFilter()
            {

                FilterCallBack2 = async () =>
                {
                    var latestData = await downloadLatest(baseUrl, datasetname);
                    if (latestData != null)
                    {
                        //discontinueIudIndicatorRow
                        var targetname = DateTime.Now;
                        var table_name = "";

                        if(Constants.ProjectTerms.RECEIVING_IUD== datasetname)
                        {
                            table_name = string.Format("lng_receive_{0}_{1}_{2}", targetname.Year, targetname.Month, targetname.Day);
                            getVmmcWebData2<receivingIudIndicatorRow>(latestData, dataMerge: 
                                new McspLngDataMergeCommand() {
                                    TargetView= "receiving",
                                    DestinationTable = table_name,
                                    datasetName = datasetname, projectName = ProjectName.MCSP });
                        }
                        else if(Constants.ProjectTerms.DISCONTINUE_IUD == datasetname)
                        {
                            table_name = string.Format("lng_discontinue_{0}_{1}_{2}",targetname.Year, targetname.Month, targetname.Day);
                            getVmmcWebData2<discontinueIudIndicatorRow>(latestData, dataMerge:                                
                                new McspLngDataMergeCommand() {
                                    TargetView= "discontinuing",
                                    DestinationTable = table_name,
                                    datasetName = datasetname, projectName = ProjectName.MCSP });
                        }
                        else
                        {
                            return;
                        }
                    }
                }
            };
            stackMain.Content = filter;
        }

        private void addQuickBooksData(object sender, RoutedEventArgs e)
        {
            var targetForm = new Forms.QuickBooksImporter()
            {
                CurrentProjectName = ProjectName.General
            };
            stackMain.Content = targetForm;
        }

        private void addTimesheetsData(object sender, RoutedEventArgs e)
        {
            var targetForm = new Forms.TimesheetImporter()
            {
                CurrentProjectName = ProjectName.General
            };
            stackMain.Content = targetForm;
        }

        private void addMaxzamData(object sender, RoutedEventArgs e)
        {
            var targetForm = new Forms.MaxzamImporter()
            {
                CurrentProjectName = ProjectName.Maxzam
            };
            stackMain.Content = targetForm;
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            Application.Current.Shutdown();
        }
    }
}
