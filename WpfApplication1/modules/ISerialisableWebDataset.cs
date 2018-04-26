using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZambiaDataManager.modules
{
    public interface ISerialisableWebDataset
    {
        DataTable getTable();
        DataRow toRow(DataRow row);
    }
}
