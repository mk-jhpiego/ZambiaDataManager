using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace ZambiaDataManager
{

    public static class MyExtensions
    {
        public static DataSet ToDataset(this List<MatchedDataValue> dataValues)
        {
            //convert to dataset
            var serialised = Newtonsoft.Json.JsonConvert.SerializeObject(dataValues);
            var table = Newtonsoft.Json.JsonConvert.DeserializeObject<DataTable>(serialised);
            var ds = new DataSet();
            ds.Tables.Add(table);
            //Assert that all require columns are available
            var expectedFields = new List<string>() { "ReportYear", "ReportMonth", "AgeGroup", "IndicatorId", "TotalCost", "OfficeAllocation" };
            foreach(var field in expectedFields)
            {
                if (!table.Columns.Contains(field))
                {
                    throw new ArgumentNullException("No column named " + field);
                }
            }
            return ds;            
        }
        public static DataSet ToDataset(this List<DataValue> dataValues)
        {
            //convert to dataset
            var ds = new DataSet();
            var table = new System.Data.DataTable() { TableName = "DataValue" };
            table.Columns.Add("FacilityName", typeof(string));
            table.Columns.Add("ReportYear", typeof(int));
            table.Columns.Add("ReportMonth", typeof(string));

            table.Columns.Add("ProgramArea", typeof(string));
            table.Columns.Add("IndicatorId", typeof(string));
            table.Columns.Add("AgeGroup", typeof(string));
            table.Columns.Add("Sex", typeof(string));
            table.Columns.Add("IndicatorValue", typeof(double));

            ds.Tables.Add(table);

            foreach (var datavalue in dataValues)
            {
                table.Rows.Add(
                    datavalue.FacilityName,
                    datavalue.ReportYear,
                    datavalue.ReportMonth,
                    datavalue.ProgramArea,
                    datavalue.IndicatorId,
                    datavalue.AgeGroup,
                    datavalue.Sex,
                    datavalue.IndicatorValue
                    );
            }
            table.AcceptChanges();
            return ds;
        }

        public static double ToDouble(this string strValue)
        {
            var asLower = strValue.ToLowerInvariant();
            if (asLower == "x" || asLower == "na")
                return Constants.NOVALUE;
            return double.Parse(asLower);
        }

        public static string csvDelim(this string strValue, bool force = false)
        {
            if (force)
                return "\"" + strValue + "\"";
            return strValue.Contains(",") ? "\"" + strValue + "\"" : strValue;
        }

        delegate void SetLabelText(string text);
        delegate void SetControlValue(int text);
        delegate void EnableControlDelegate(bool stateToSet);
        public static void SetText(this Control control, string text)
        {
            //if (control.InvokeRequired)
            //{
            //    control.Invoke(new SetLabelText((txt) => { control.Text = txt; }), text);
            //    return;
            //}


            //control.SetValue()

           // control.SetText(text);
        }
        public static void SetValue(this ProgressBar control, int text)
        {
            //if (control.InvokeRequired)
            //{
            //    control.Invoke(new SetControlValue((txt) => { control.Value = txt; }), text);
            //    return;
            //}
            control.Value = text;
        }

        public static void SetStepValue(this ProgressBar control, int text)
        {
            control.SetValue(text);
            //control.SetStepValue(text);
        }

        public static void EnableControl(this Control control, bool newState)
        {
            //if (control.InvokeRequired)
            //{
            //    control.Invoke(new EnableControlDelegate((txt) => { control.Enabled = txt; }), newState);
            //    return;
            //}
            control.IsEnabled = newState;
        }
    }

    //public class DummyControl :Control
    //{
    //    public bool Enabled { get; set; }

    //    public bool InvokeRequired { get; set; }

    //    internal void Invoke(object enableControlDelegate, bool newState)
    //    {
    //        //throw new NotImplementedException();
    //    }
    //}

}
