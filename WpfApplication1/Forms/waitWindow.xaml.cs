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
    public partial class waitWindow : Page
    {
        public waitWindow()
        {
            InitializeComponent();
        }

        private bool _showWait;
        public bool showWait
        {
            get { return _showWait; }
            set
            {
                _showWait = value;
                if (!_showWait)
                {
                    lblMainMessage.Content = "Please wait. App is getting started, ...";
                }
                else
                {
                    lblMainMessage.Content = "Ready";
                }
            }
        }

        public string displayMsg
        {
            get { return ""; }
            set
            {
                if (!string.IsNullOrWhiteSpace(value))
                {
                    lblMainMessage.Content = value;
                }
            }
        }
    }
}
