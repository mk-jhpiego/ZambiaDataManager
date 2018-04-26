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
using System.Windows.Shapes;

namespace ZambiaDataManager.Popups
{
    /// <summary>
    /// Interaction logic for SplashScreen.xaml
    /// </summary>
    public partial class ProjectedSelector : Window
    {
        public ProjectedSelector()
        {
            InitializeComponent();
            SelectedProjectName = ProjectName.None;
        }

        public ProjectName SelectedProjectName { get; set; }

        private void bVmmc_Click(object sender, RoutedEventArgs e)
        {
            var asButton = sender as Button;
            if (asButton == null)
                return;
            var name = asButton.Name;
            if (name == "bDOD")
            {
                SelectedProjectName = ProjectName.DOD;
            }
            else if (name == "bVmmc")
            {
                SelectedProjectName = ProjectName.IHP_VMMC;
            }
            else if (name == "bMcsp")
            {
                SelectedProjectName = ProjectName.MCSP;
            }
            else if (name == "bTraining")
            {
                SelectedProjectName = ProjectName.IHP_Capacity_Building_and_Training;
            }
            else if (name == "bGeneral")
            {
                SelectedProjectName = ProjectName.General;
            }
            this.DialogResult = true;
        }

        public bool RememberSelection { get { return chkRememberSelection.IsChecked ?? false; } }
    }
}
