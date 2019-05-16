using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DODExcelReader
{
    class Program
    {
        static void Main(string[] args)
        {
            //var getValues = true;
            Console.WriteLine("Press any key to start");
            var g = Console.ReadLine();
            var ignoredFIles =  File.ReadAllLines("toIgnore.txt");            
            File.AppendAllText("headermatches.csv", "header,rowid,columnidd,districtname,filename\r\n");

            Microsoft.Office.Interop.Excel.Application excelapp = new Microsoft.Office.Interop.Excel.Application() { Visible = false };

            var folder = @"C:\Users\makando\Desktop\DOD Monthly";
            var directory = Directory.GetFiles(folder, "*.xls", SearchOption.AllDirectories);
            Console.WriteLine(string.Format("Files found: {0}", directory.Length));
            var cntr = 0;
            directory.ToList().ForEach(x =>
            {
                if (ignoredFIles.Contains(x)) return;

                //if (cntr >= 3) return;

                Console.WriteLine("Processing {0}", x);
                var excel = new GetExcelAsDataTable2()
                {
                    excelApp = excelapp,
                    addHeaderRow = false, // cntr == 0,
                    fileName = x,
                    rootPath= folder
                };
                //var t = excel.Execute();
                var t = excel.ExecuteDataset();
                if (t == null)
                {

                }
                else
                {
                    //we save to the database
                    Console.WriteLine("Saving to database");
                    new SaveTableToDbCommand()
                    {
                        TargetDataset = t
                    }.Execute();

                    Console.WriteLine("Merging data");
                    var dataMerge = new TempDodMergeCommand()
                    {
                        tempName = t.Tables[0].TableName
                    };

                    // we save, 
                    dataMerge.DoMerge();
                    //we drop the table
                    Console.WriteLine("Data merged");

                    //File.AppendAllText(@"ds\compiledfiles.txt", x + "\r\n");
                    //foreach (var key in t.Keys)
                    //{
                    //    var lst = t[key];
                    //    File.AppendAllLines(@"ds\"+key + "-ds.csv", lst);
                    //}
                    Console.WriteLine(cntr++);
                }
                
            });

            Console.WriteLine("Any key to exit");
            Console.ReadLine();
        }
    }
}
