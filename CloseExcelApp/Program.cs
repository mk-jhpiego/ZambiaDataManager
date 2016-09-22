using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KillAllExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            var processes = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            processes.ToList().ForEach(t => t.Kill());
        }
    }
}
