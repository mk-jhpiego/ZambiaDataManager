using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
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
using ZambiaDataManager.runners;

namespace ZambiaDataManager
{
    /// <summary>
    /// Interaction logic for datamanager.xaml
    /// </summary>
    public partial class datamanager : Window
    {
        public datamanager()
        {
            InitializeComponent();
        }

        async Task<byte[]> getDocument(string url)
        {
            await Task.Delay(10000);
            return new byte[3];
            //var client = new HttpClient()
            //{
            //    BaseAddress = new Uri(url)
            //};
            //return await client.GetByteArrayAsync(url);
        }

        void doStartup()
        {

        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            var form = new Popups.Splash();
            form.Show();
            var url = "https://www.dropbox.com/s/ilp2atvhy0mvohj/Zambia%20DHS%202013%20Final%20with%20cover.pdf?dl=0";
            //var t = await getDocument(url);
            var status = await mainapp.Instance.initialise(
                onSuccess: null,
                onFailure: null,
                onNoConnection: null,
                waitScreen: null
                );
            form.Close();
            var mainWindow = new MainWindow();
            mainWindow.Show();
            this.Hide();
            //we show the main document
            //var initialiser = new appInitialiser() {
            //    Code = async () => {
            //        var url = "https://www.dropbox.com/s/ilp2atvhy0mvohj/Zambia%20DHS%202013%20Final%20with%20cover.pdf?dl=0";
            //        return await getDocument(url).Result;
            //    }
            //};
            //var fResult = initialiser.Execute().flip();
            //if (fResult == resultStatus.startupSuccessful)
            //{
            //    startupSuccessful("");
            //}
            //else if (fResult == resultStatus.startupFailed)
            //{
            //    startupFailed("");
            //}
            //else if (fResult == resultStatus.startupNoConnection)
            //{
            //    startupNoConnection("");
            //}
            ////we show the main window
        }

        async Task<resultStatus> initialiseApp(IDisplayProgress splash)
        {
            return await mainapp.Instance.initialise(
                onSuccess: null,
                onFailure: null,
                onNoConnection: null,
                waitScreen: null
                );
        }

        void closeWaitScreen()
        {
            //try
            //{
            //    if (_waitScreen != null)
            //    {
            //        _waitScreen.Close();
            //        _waitScreen = null;
            //    }
            //}
            //catch (Exception c)
            //{

            //}
        }

        void startupNoConnection(string msgAction)
        {
            closeWaitScreen();
        }
        void startupFailed(string msgAction)
        {
            closeWaitScreen();
        }
        void startupSuccessful(string msgAction)
        {
            closeWaitScreen();
            Task.Run(() =>
            {
                var mainWindow = new MainWindow();
                mainWindow.Show();
            }            
            );
        }
    }
}
