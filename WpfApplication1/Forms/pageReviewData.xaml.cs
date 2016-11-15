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
    public partial class pageReviewData : Page
    {
        DbHelper dbhelper = null;
        DataTable currentList = null;
        public pageReviewData()
        {
            InitializeComponent();
            var dbbuilder = DbFactory.GetDefaultConnection(PageController.Instance.DefaultProjectName);
            if (dbbuilder == null)
            {
                //perhaps we fail to load and throw an exception
            }
            dbhelper = new DbHelper(dbbuilder);

            currentList = fetchData();
            refreshDataGrid(currentList);
        }

        private void deleteSelectedRow(object sender, RoutedEventArgs e)
        {
            if (currentList == null || gIntermediateData.SelectedItem == null)
                return;

            var currentRow = gIntermediateData.SelectedItem as DataRowView;

            //we get the correct details from  the server
            var facilityindex = dbhelper.GetScalar("select FacilityIndex From FacilityList where DistrictName = @DistrictName and FacilityName = @FacilityName",
                new CommandParam().Add("DistrictName", currentRow["DistrictName"])
                .Add("FacilityName", currentRow["FacilityName"])
                );

            var year = dbhelper.GetScalar("select YearId From YearLookup where YearName = @ReportYear",
    new CommandParam().Add("ReportYear", currentRow["ReportYear"]));

            var month = dbhelper.GetScalar("select MonthID From MonthLookUp where MonthName = @ReportMonth",
new CommandParam().Add("ReportMonth", currentRow["ReportMonth"]));

            //delete the selected row
            var user = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            dbhelper.ExecProc("dbo.proc_BackupAndDelete",
                new CommandParam().Add("FacilityIndex", facilityindex)
                .Add("ReportYear", year)
                .Add("ReportMonth", month)
                .Add("username", user)
                );

            currentList = fetchData();
            refreshDataGrid(currentList);
        }

        DataTable fetchData()
        {
            var sql =
                //@"SELECT distinct 
                // fl.ProvinceName, fl.DistrictName, fl.FacilityName, 
                //convert(varchar,yr.[YearName])+' '+
                //substring(mnth.[MonthName],1,3)
                //AS 'YearMonth',
                //yr.YearID, 
                //mnth.MonthID,
                // pa.ProgramArea, 
                //f.FacilityIndex,
                //pa.ProgramAreaID,
                //yr.YearName AS ReportYear, 
                //mnth.[MonthName] AS ReportMonth
                //FROM FacilityData  f 
                //join YearLookUp yr on f.ReferenceYear = yr.YearID
                //join MonthLookUp mnth on f.ReferenceMonth = mnth.MonthID
                //join FacilityList fl on f.FacilityIndex = fl.FacilityIndex 
                //join IndicatorLookup il on f.IndicatorSerial = il.IndicatorSerial
                //join ProgramAreaLookUp pa on il.ProgramAreaID = pa.ProgramAreaID
                //";
                "select distinct ProvinceName, DistrictName, FacilityName, [Year] as ReportYear, [Month] as ReportMonth From MainData";

            return dbhelper.GetTable(sql);
        }

        public List<FileDetails> SelectedFiles { get; private set; }

        void refreshDataGrid(DataTable dataSource = null)
        {
            gIntermediateData.ItemsSource = "";
            gIntermediateData.ItemsSource = dataSource.DefaultView;
        }

        private void openDataset(object sender, RoutedEventArgs e)
        {
            fetchData();
        }
    }
}
