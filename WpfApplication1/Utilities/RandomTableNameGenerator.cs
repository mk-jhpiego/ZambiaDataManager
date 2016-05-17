using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZambiaDataManager.Utilities
{
    class RandomTableNameGenerator : IQueryHelper<string>
    {
        public IDisplayProgress progressDisplayHelper { get; set; }

        public string Execute()
        {
            return "etl_" + Path.GetFileNameWithoutExtension(Path.GetRandomFileName());
        }
    }
}
