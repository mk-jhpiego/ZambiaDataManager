using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ZambiaDataManager.Popups;

namespace ZambiaDataManager.runners
{
    public class appInitialiser : IQueryHelper<resultStatusClass>
    {
        public resultStatusClass Execute()
        {
            if (Code != null)
                return Code().flip();
            return resultStatus.startupFailed.flip();
        }
        public Func<resultStatus> Code { get; set; }
        public IDisplayProgress progressDisplayHelper { get; set; }
        public Action<string> Alert { get; set; }
    }

    //public class CodeRunner<T> : ICommandExecutor where T : class
    //{
    //    public IQueryHelper<T> CodeToExcute { get; internal set; }

    //    public Action<T> AsyncCallBack { private get; set; }

    //    delegate void closeFormDelegate();

    //    Splash defaultSplashScreen = null;
    //    public bool ShowSplash { get; internal set; }
    //    void closeForm()
    //    {
    //        defaultSplashScreen.Close();
    //        //defaultSplashScreen.Dispose();
    //    }

    //    void closeSplash()
    //    {
    //        closeForm();
    //    }

    //    public void Execute()
    //    {
    //        if (CodeToExcute == null)
    //        {
    //            return;
    //        }

    //        defaultSplashScreen = new Splash() {
    //            WindowStartupLocation = 
    //            System.Windows.WindowStartupLocation.CenterOwner };
    //        CodeToExcute.progressDisplayHelper = defaultSplashScreen;
    //        var task = new Task(() =>
    //        {
    //            var res = CodeToExcute.Execute();
    //            closeSplash();
    //            if (AsyncCallBack != null)
    //            {
    //                AsyncCallBack(res);
    //            }
    //        });
    //        task.Start();
    //        defaultSplashScreen.ShowDialog();
    //    }
    //}
}
