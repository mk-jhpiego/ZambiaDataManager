using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace ZambiaDataManager
{
    public class ParamPair
    {
        public ParamPair(string name, object value)
        {
            Name = name;
            Value = value;
        }
        public string Name;
        public object Value;
    }

    public class CommandParam
    {
        public List<ParamPair> Parameters = new List<ParamPair>();
        public CommandParam Add(string name, object value)
        {
            Parameters.Add(new ParamPair(name, value));
            return this;
        }
    }

    public interface ICommandExecutor
    {
        void Execute();
    }
    public enum ProjectName
    {
        None = 1,
        DOD = 2,
        IHP_VMMC,
        IHP_Capacity_Building_and_Training,
        General,
        Finance
    }

    public class FileDetails
    {
        public string FileName { get; set; }
    }

    public class FinanceDetails:FileDetails
    {
        public string ReportMonth { get; set; }
        public int ReportYear { get; set; }
    }


    public interface IQueryHelper<T> where T : class
    {
        T Execute();
        IDisplayProgress progressDisplayHelper { get; set; }

        Action<string> Alert { get; set; }
    }

    public interface IDisplayProgress
    {
        void PerformProgressStep(string message = "");
        void MarkStartOfMultipleSteps(int stepsToExpect);
        void ResetSubProgressIndicator(int stepsToExpect);
        void PerformSubProgressStep();
    }

    //public interface ICommandHelper<T> where T : class
    //{
    //    void Execute();
    //}

    public class LocationDetail
    {
        public string FacilityName { get; set; }
        public string ReportMonth { get; set; }
        public int ReportYear { get; internal set; }
        //public string ReportYear { get; set; }
    }

    public class TwoDataValuePair
    {
        public DataValue TotalCostDataValue = null, OfficeAllocationDataValue = null;
        public MatchedDataValue AsMatchedDataValue(LocationDetail location)
        {
            var sourceObj = TotalCostDataValue ?? OfficeAllocationDataValue;
            var t=new MatchedDataValue()
            {
                IndicatorId = sourceObj.IndicatorId,
                AgeGroup = sourceObj.AgeGroup,

                OfficeAllocation = OfficeAllocationDataValue == null ? 0: OfficeAllocationDataValue.IndicatorValue,
                TotalCost = TotalCostDataValue == null ? 0 : TotalCostDataValue.IndicatorValue,
                ReportMonth = location.ReportMonth,
                ReportYear = location.ReportYear,
                FacilityName = location.FacilityName
            };
            t.DirectCost = t.TotalCost - t.OfficeAllocation;
            
            return t;
        }
    }

    public class MatchedDataValue: DataValue
    {
        public double OfficeAllocation { get; set; }
        public double TotalCost { get; set; }
        public double DirectCost { get; set; }
        public string ProjectMatchKey { get; internal set; }
    }

    public class DataValue
    {
        public string FacilityName { get; set; }
        public int ReportYear { get; set; }
        public string ReportMonth { get; set; }

        public string ProgramArea { get; set; }
        public string IndicatorId { get; set; }
        public double IndicatorValue { get; set; }
        public string AgeGroup { get; internal set; }
        public string Sex { get; internal set; }
    }

    public class ProgramAreaDefinition
    {
        public ProgramAreaDefinition()
        {
            ProgramArea = string.Empty;
            //ServiceAreas = new ServiceAreaDataset();
            Indicators = new List<ProgramIndicator>();
            AgeDisaggregations = new List<string>();

            DefaultHandler = "default";
            Gender = string.Empty;
        }

        //public ServiceAreaDataset ServiceAreas { get; set; }
        public string ProgramArea { get; set; }
        public List<ProgramIndicator> Indicators { get; set; }
        public List<string> AgeDisaggregations { get; set; }

        public string DefaultHandler { get; set; }
        public string Gender { get; set; }
    }

    public class ServiceAreaDataset
    {
        public ServiceAreaDataset()
        {
            ProgramArea = string.Empty;
            AgeDisaggregations = new List<string>();
            DefaultHandler = string.Empty;
            Gender = string.Empty;
        }

        public string ProgramArea { get; set; }
        public List<string> AgeDisaggregations { get; set; }
        public string DefaultHandler { get; set; }
        public string Gender { get; set; }
    }

    public class ProgramIndicator
    {
        public ProgramIndicator()
        {
            IndicatorId = string.Empty;
            Indicator = string.Empty;
        }

        public string IndicatorId { get; set; }
        public string Indicator { get; set; }
        public List<string> SubAgeDisaggregations { get; set; }
    }

    public class ProgramAreaIndicators
    {
        public ProgramAreaIndicators()
        {
            ProgramArea = string.Empty;
            Indicators = new List<ProgramIndicator>();
        }

        public string ProgramArea { get; set; }
        public List<ProgramIndicator> Indicators { get; set; }
    }

    public class RowColmnPair
    {
        public RowColmnPair(int rowId, int colmn1, int colmn2)
        {
            Row = rowId;
            Column = colmn1;
            Column2 = colmn2;
        }

        public RowColmnPair()
        {
        }
        public int Row { get; set; }
        public int Column { get; set; }
        public int Column2 { get; set; }
    }
}
