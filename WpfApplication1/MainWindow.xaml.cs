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

namespace WpfApplication1
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

        }

        private void addVmmcCampaignDailyData(object sender, RoutedEventArgs e)
        {

        }

        private void reviewUploadedData(object sender, RoutedEventArgs e)
        {
            var form = new Forms.pageAddMERLData();
            this.stackMain.Children.Clear();
            this.stackMain.Children.Add(form);
            //we create an instance of d:\my documents\visual studio 2015\Projects\WpfApplication1\WpfApplication1\Forms\pageAddMERLData.xaml
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
    }
}
