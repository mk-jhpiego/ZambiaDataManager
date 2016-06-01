using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZambiaDataManager
{
    public class Constants
    {
        public const double NOVALUE = -999999;
        public const string OFFICE_ALLOCATION = "office allocation";
        public const string INCOUNTRY_ION_EXPENSES = "Incountry IONs expenses";
        public const string TOTAL_EXPENDITURE = "total_expenditure";

        public static List<string> monthsLongName = new List<string>() { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };
        public static List<string> monthsShortName = new List<string>() { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
        public static List<string> ihpMonthNames = new List<string>() { "Jan_1", "Feb_2", "Mar_3", "Apr_4", "May_5", "Jun_6", "Jul_7", "Aug_8", "Sep_9", "Oct_10", "Nov_11", "Dec_12" };

        public static List<int> acceptableYears = new List<int>() { 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018 };


        public static string GetStandardMonthName(string userMonthName)
        {
            var lower = userMonthName.ToLowerInvariant();
            var monthName = string.Empty;
            if (!monthsLongName.Select(t => t.ToLowerInvariant()).Contains(lower))
            {
                //we check the short form
                if (lower == "sept")
                {
                    lower = "sep".ToLowerInvariant();
                }
                else if (lower.Length != 3 || !monthsShortName.Select(t => t.ToLowerInvariant()).Contains(lower))
                {
                    monthName = null;
                }
                var monthIndx = monthsShortName.FindIndex(t => t.ToLowerInvariant() == lower);
                monthName = monthsLongName[monthIndx];
            }
            else
            {
                var monthIndx = monthsLongName.FindIndex(t => t.ToLowerInvariant() == lower);
                monthName = monthsLongName[monthIndx];
            }
            return monthName;
        }

        internal static string GetAlternateStandardMonthName(string mValue)
        {
            var monthName = string.Empty;
            if (!string.IsNullOrWhiteSpace(mValue) && ihpMonthNames.Contains(mValue.Trim()))
            {
                var lower = mValue.Substring(0, 3).ToLowerInvariant();
                var mIndex = monthsShortName.FindIndex(t => t.ToLowerInvariant() == lower);
                monthName = monthsLongName[mIndex];
            }
            else
            {
                monthName = null;
            }
            return monthName;
        }
    }
}
