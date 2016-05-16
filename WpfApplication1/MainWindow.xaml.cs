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
            var targetForm = new Forms.pageAddMERLData();
            stackMain.Content = targetForm;
        }

        private void addVmmcCampaignDailyData(object sender, RoutedEventArgs e)
        {

        }

        private void reviewUploadedData(object sender, RoutedEventArgs e)
        {

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
    }
}
