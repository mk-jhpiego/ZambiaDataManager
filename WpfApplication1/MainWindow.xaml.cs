using System;
using System.Collections.Generic;
using System.Data.SqlClient;
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

        }

        private void showPepfarReport(object sender, RoutedEventArgs e)
        {

        }

        private void showAdminSummaryReport(object sender, RoutedEventArgs e)
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

        private void bViewPepfarReport_Click(object sender, RoutedEventArgs e)
        {

        }

        void setMainMenuStatus(Visibility status)
        {
            stackUserMenu.Visibility = status;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.Hide();
            Application.Current.DispatcherUnhandledException += Current_DispatcherUnhandledException;

            //we show the wait window
            setMainMenuStatus(Visibility.Hidden);

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

            //we load db data
            var shutDown = false;

            //we check the connection
            if (!isConnected())
                return;

            try
            {             
                loadDbData();
                setMainMenuStatus(Visibility.Visible);
            }
            catch (SqlException sqlex)
            {
                MessageBox.Show("Could not connect to the database. The application will shut down");
                shutDown = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not start the application. Error has been logged. The application will shut down");
                shutDown = true;
            }
            finally
            {
            }
            if (shutDown)
            {
                Application.Current.Shutdown();
            }
            this.Show();
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
                var test = db.GetScalar("select 1", 3);
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
