using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace ZambiaDataManager.Forms
{
    /// <summary>
    /// Interaction logic for pageAddMERLData.xaml
    /// </summary>
    public partial class pageYearMonthFilter : Page
    {
        public pageYearMonthFilter()
        {
            InitializeComponent();
        }

        public string procedureName { get; set; }
        public Action<int, string> FilterCallBack { get; internal set; }
        public Action FilterCallBack2 { get; internal set; }


        public Func<string, List<KeyValuePair<string, object>>, DataTable> refreshHandler;

        public void refreshData(object sender, RoutedEventArgs ev)
        {
            labelStatus.Content = "...";
            if (FilterCallBack2 != null)
            {
                FilterCallBack2();
            }
            else
            {
                //var yearMonth = textYearMonth.Text;
                //if (refreshHandler != null && !string.IsNullOrWhiteSpace(yearMonth))
                //{
                //    var table = refreshHandler(procedureName,
                //        new List<KeyValuePair<string, object>>() {
                //            new KeyValuePair<string, object>( "@yearMonth",yearMonth)
                //        }
                //        );
                //    refreshDataGrid(table);
                //}
                int selectedYear = -1;
                var currentYr = DateTime.Now.Year;
                if (!int.TryParse(textYearMonth.Text, out selectedYear) || selectedYear < 2016 || selectedYear > currentYr)
                {
                    labelStatus.Content = "Please specify a correct year";
                    return;
                }

                var selectedMonth = panelMonths.Children.OfType<RadioButton>()
                    .FirstOrDefault(t => t.IsChecked == true);
                if (selectedMonth == null)
                {
                    labelStatus.Content = "Please select a month";
                    return;
                }
                //we convert to int
                var txt = Convert.ToString(selectedMonth.Content);
                if (txt.ToLowerInvariant() == "sept")
                {
                    txt = "Sep";
                }
                var indx = Constants.monthsShortName.IndexOf(txt);
                if (indx == -1)
                {
                    throw new ArgumentOutOfRangeException("Could not determine the month for " + txt);
                }

                indx += 1;

                FilterCallBack(selectedYear, (indx < 10 ? "0" + indx : "" + indx));
            }
        }

        private void textYearMonth_TextChanged(object sender, TextChangedEventArgs e)
        {
            var txt = e.Source as TextBox;
            var val = txt.Text;
            if (string.IsNullOrWhiteSpace(val))
                return;
            int v;
            if ( !int.TryParse(val, out v))
            {
                txt.Clear();
                e.Handled = true;
            }
        }
    }
}
