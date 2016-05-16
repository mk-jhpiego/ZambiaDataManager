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

namespace WpfApplication1.Forms
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

        List<string> _selectedFiles = new List<string>();

        public List<string> SelectedFiles
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
            //using(var dialog = new FileSelectDialog())
            //{

            //}
        }


    }
}
