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
        }

        List<FileDetails> _selectedFiles = new List<FileDetails>();

        public List<FileDetails> SelectedFiles
        {
            get
            {
                return _selectedFiles;
            }

            set
            {
                _selectedFiles = value;
            }
        }

        private void selectFile(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog()
            {
                CheckFileExists = true,
                Multiselect = true,
                CheckPathExists = true,
                Filter = "Excel (*.xlsx)|*.xlsx",
                Title = "Please select the files to import"
            };
            var dialogResult = dialog.ShowDialog() ?? false;
            if (dialogResult)
            {
                _selectedFiles.AddRange(
                    (from file in dialog.FileNames
                     select new FileDetails() { FileName = file }).ToList()
                    );
            }

            //we refresh the grid
            refreshDataGrid();
        }

        void refreshDataGrid()
        {
            if (gSelectedFiles.ItemsSource == null)
            {
                gSelectedFiles.ItemsSource = _selectedFiles;
            }
            else
            {
                gSelectedFiles.ItemsSource = "";
                gSelectedFiles.ItemsSource = _selectedFiles;
            }
        }

        private void uploadSelectedFiles(object sender, RoutedEventArgs e)
        {

        }

        private void clearSelected(object sender, RoutedEventArgs e)
        {

        }
    }
}
