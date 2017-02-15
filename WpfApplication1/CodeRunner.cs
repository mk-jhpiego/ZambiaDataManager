using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ZambiaDataManager.Popups;

namespace ZambiaDataManager
{
    public class CodeRunner<T> : ICommandExecutor where T : class
    {
        public IQueryHelper<T> CodeToExcute { get; internal set; }

        public Action<T> AsyncCallBack { private get; set; }

        delegate void closeFormDelegate();

        WaitDialog defaultSplashScreen = null;
        public bool ShowSplash { get; internal set; }
        void closeForm()
        {
            defaultSplashScreen.Close();
            //defaultSplashScreen.Dispose();
        }

        void closeSplash()
        {
            closeForm();
        }

        public void Execute()
        {
            if (CodeToExcute == null)
            {
                return;
            }

            defaultSplashScreen = new WaitDialog() { WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner };
            CodeToExcute.progressDisplayHelper = defaultSplashScreen;
            var task = new Task(() =>
            {
                var res = CodeToExcute.Execute();
                closeSplash();
                if (AsyncCallBack != null)
                {
                    AsyncCallBack(res);
                }
            });
            task.Start();
            defaultSplashScreen.ShowDialog();
        }
    }
}
