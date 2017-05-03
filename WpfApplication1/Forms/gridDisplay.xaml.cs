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
    /// Interaction logic for pageAddMERLData.xaml
    /// </summary>
    public partial class gridDisplay : Page
    {
        public gridDisplay()
        {
            InitializeComponent();
        }

        public string procedureName { get; set; }
        public Func<string, List<KeyValuePair<string, object>>, DataTable> refreshHandler;

        public void refreshData(object sender, RoutedEventArgs ev)
        {
            var yearMonth = textYearMonth.Text;
            if (refreshHandler != null && !string.IsNullOrWhiteSpace(yearMonth))
            {
                var table = refreshHandler(procedureName,
                    new List<KeyValuePair<string, object>>() {
                        new KeyValuePair<string, object>( "@yearMonth",yearMonth)
                    }
                    );
                refreshDataGrid(table);
            }
        }

        public void refreshDataGrid(DataTable dataSource)
        {
            gIntermediateData.ItemsSource = "";
            gIntermediateData.ItemsSource = dataSource.DefaultView;
        }
    }
}
