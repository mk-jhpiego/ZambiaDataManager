using System;
using System.Collections.Generic;
using System.IO;
using System.IO.IsolatedStorage;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using ZambiaDataManager.Storage;

namespace ZambiaDataManager
{
    public class mainapp
    {
        static mainapp _instance;
        static object _mylock = new object();
        public static mainapp Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_mylock)
                    {
                        if (_instance == null)
                            _instance = new mainapp() {Worker = new mainAppWorker() };
                    }
                } 
                return _instance;
            }
        }

        public mainAppWorker Worker { get; set; }

        internal void showView(System.Windows.Controls.Page page, ProjectName CurrentProjectName)
        {

        }

        async internal Task<resultStatus> initialise(IDisplayProgress waitScreen, Action<string> onFailure, 
            Action<string> onSuccess, Action<string> onNoConnection)
        {
            var then = DateTime.Now;
            var url = "https://www.dropbox.com/s/ilp2atvhy0mvohj/Zambia%20DHS%202013%20Final%20with%20cover.pdf?dl=0";
            var resource = await getDocument(url);
            var size = resource.Length;
            //await Task.Delay(5000);
            var now = DateTime.Now;
            var diff = now.Subtract(then).TotalMilliseconds;
            //onSuccess("" + diff);
            return resultStatus.startupSuccessful;
        }

        async Task<byte[]> getDocument(string url)
        {
            var client = new HttpClient() {
                BaseAddress = new Uri(url)
            };
            return await client.GetByteArrayAsync(url);
        }
    }

    public class mainAppWorker
    {
        PreferencesHelper _preferences;
        ServerPasswordsHelper _ServerPasswordsHelper;
        public mainAppWorker()
        {
            _preferences = new PreferencesHelper(
                getAppFolder(Constants.CommonFolders.WORKINGFOLDER));

            _ServerPasswordsHelper = new ServerPasswordsHelper();
            var uname = _ServerPasswordsHelper.get("username");
            var pwd = _ServerPasswordsHelper.get("password");
            var servername = _ServerPasswordsHelper.get("server");

            DbFactory.ServerName = servername;
            DbFactory.InstanceName = "";
            DbFactory.username = uname;
            DbFactory.password = pwd;

            //a dirty catch to avoid messing with the server. Feel free to remove
            if (Environment.MachineName == "D-9W48GC2"
                || Environment.MachineName == "SUPER-LAP")
            {
                var res = MessageBox.Show("Use your Local Computer rather than the server MK ???????????", "WAIT!!!!!!!!!!!", MessageBoxButton.OKCancel);
                if (res == MessageBoxResult.OK)
                {
                    if (Environment.MachineName == "D-9W48GC2")
                    {
                        DbFactory.ServerName = "D-9W48GC2";
                        DbFactory.InstanceName = "SQLDEV";
                    }
                    else
                    {
                        DbFactory.ServerName = "SUPER-LAP";
                        DbFactory.InstanceName = "SQL2014";
                    }
                }
            }
        }

        string checkfolder(string folderPath)
        {
            if (!Directory.Exists(folderPath))
            {
                try
                {
                    Directory.CreateDirectory(folderPath);
                }
                catch (IOException ioex)
                {
                    throw;// new Exception("Error creating directory. IOExce");
                }
                catch (Exception ex)
                {
                    throw;
                }
            }
            return folderPath;
        }

        public void setDefaultProject(ProjectName project)
        {
            _preferences.set(Constants.Preferences.PROJECT, 
                Enum.GetName(typeof(ProjectName), project));
        }

        public ProjectName getDefaultProject()
        {
            var defaultproject = ProjectName.None;
            var userPreference = _preferences.get(Constants.Preferences.PROJECT);
            if (!string.IsNullOrWhiteSpace(userPreference))
            {
                if (!Enum.TryParse(userPreference, out defaultproject))
                {
                    defaultproject = ProjectName.None;
                }
            }
            return defaultproject;
        }

        string getAppFolder(string folderName)
        {
            var userDirectory = Environment.GetFolderPath(
                Environment.SpecialFolder.MyDocuments);
            var path = Path.Combine(userDirectory, folderName);
            return checkfolder(path);
        }
    }

    public class ServerPasswordsHelper : PropertyStorageHelper
    {
        public ServerPasswordsHelper() : base("DbData\\dbaccess.json")
        {

        }
    }
    public class PreferencesHelper: PropertyStorageHelper
    {
        public PreferencesHelper(string workingFolder):base(Path.Combine(workingFolder,
                Constants.CommonFolders.USEROPTIONS))
        {

        }
    }

    public class PropertyStorageHelper
    {
        //string appFolder;
        protected string userOptionsFilePath;
        public PropertyStorageHelper(string workingFolder)
        {
            userOptionsFilePath = workingFolder;
        }

        public void set(string preference, string value)
        {
            var prefs = getUserPreferences();
            if (prefs == null)
                prefs = new Dictionary<string, string>();

            prefs[preference] = value;
            File.WriteAllText(userOptionsFilePath,
                Newtonsoft.Json.JsonConvert.SerializeObject(prefs));
        }

        public string get(string preference)
        {
            var toReturn = string.Empty;
            var userPreferences = getUserPreferences();
            if (userPreferences != null && userPreferences.ContainsKey(preference))
            {
                toReturn = userPreferences[preference];
            }
            return toReturn;
        }

        private Dictionary<string, string> getUserPreferences()
        {
            Dictionary<string, string> toReturn = null;
            //we check if we have the file for settings
            if (File.Exists(userOptionsFilePath))
            {
                var filecontents = File.ReadAllText(userOptionsFilePath);
                var deserialised = Newtonsoft.Json.JsonConvert
                    .DeserializeObject<Dictionary<string, string>>(filecontents);
                toReturn = deserialised;
            }
            return toReturn;
        }
    }

    public enum resultStatus
    {
        startupSuccessful, startupFailed, startupNoConnection
    }

    public class resultStatusClass
    {
        resultStatus _carried;
        public resultStatusClass(resultStatus status) {
            ResultStatus = status;
        }

        public resultStatus ResultStatus { get; set; }
    }
}
