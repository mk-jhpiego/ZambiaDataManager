using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ZambiaDataManager.Storage
{
    public class AgegroupsProvider
    {
        public DbHelper DB { get; set; }
        const string ageGroupSql = "select a.AgeGroup AlterNameAgeGroup, l.AgeGroupName StandardAgeGroup From AgeGroupLookupAlternate a join AgeGroupLookUp l on a.AgeGroupId = l.AgeGroupID";
        public Dictionary<string,string> getAlternateAgeGroups()
        {
            //GetLookups
            var toReturn = new Dictionary<string, string>();
            var lookups = DB.GetLookups(ageGroupSql);
            lookups.ToList().ForEach(lookup => toReturn.Add(Convert.ToString(lookup.Key), Convert.ToString(lookup.Value)));
            //from lookup in lookups
            //let key = Convert.ToString(lookup.Key)
            //let value = Convert.ToString(lookup.Value)
            return toReturn;
        }
    }

    public class BaseMergeCommand : IQueryHelper<IEnumerable<string>>
    {
        public ProjectName projectName { get; set; }
        public string datasetName { get; set; }
        public bool IsWebData { get; set; }
        public Action<string> Alert { get; set; }
        public DbHelper Db { get; set; }

        public bool IsInError { get; set; }
        public string DestinationTable { get; set; }
        public string TempTableName { get; set; }

        public IDisplayProgress progressDisplayHelper { get; set; }
        public IEnumerable<string> Execute()
        {
            //
            DoMerge();
            return new List<string>();
        }

        protected virtual void DoMerge()
        {
        }
    }
}
