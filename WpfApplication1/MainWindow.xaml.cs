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

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Application.Current.DispatcherUnhandledException += Current_DispatcherUnhandledException;
            //we prompt for the default database
            var dialog = new ProjectedSelector();
            if (dialog.ShowDialog() != null && dialog.SelectedProjectName != ProjectName.None)
            {
                PageController.Instance.DefaultProjectName = dialog.SelectedProjectName;
                Title = "Jhpiego Zambia Data Manager: Default project selected is " + dialog.SelectedProjectName.ToString();
                var user = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                tLoggedInUser.Text = user??"Not Logged In" ;

                //we load db data
                var shutDown = false;
                try
                {
                    loadDbData();
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
            }
            else
            {
                //means they canceled
                //we disable everything
                Application.Current.Shutdown();
            }
        }

        private void Current_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            MessageBox.Show("An issue is preventing startup. Please retry. The app will shutdown. If the problem persists, please call for support");
            e.Handled = true;
            Application.Current.Shutdown();
        }

        void loadDbData()
        {
            var db = new DbHelper(DbFactory.GetDefaultConnection(PageController.Instance.DefaultProjectName));
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
