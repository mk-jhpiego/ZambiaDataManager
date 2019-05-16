using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HPVExcelReader
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
            
            var folder = @"K:\development\github\ZambiaDataManager\HPVExcelReader\HPVExcelReader\bin\hpvdocs";
            var directory = Directory.GetFiles(folder, "*.*", SearchOption.AllDirectories);
            Console.WriteLine(string.Format("Files found: {0}", directory.Length));
            var cntr = 0;
            var filecounter = 0;

            var processingMicroplans = false;
            if (processingMicroplans)
            {
                directory.ToList().ForEach(x =>
                {
                    Console.WriteLine(filecounter++);
                    if (ignoredFIles.Contains(x))
                    {
                        File.AppendAllText(@"ds\files_ingnored.txt", x + "\r\n");
                        return;
                    }

                    //if (cntr >= 30) return;

                    var excel = new GetExcelAsDataTable2()
                    {
                        excelApp = excelapp,
                        addHeaderRow = cntr == 0,
                        fileName = x
                    };
                    var t = excel.Execute();
                    if (t == null)
                    {
                        Console.WriteLine("Not compiled", x);
                        File.AppendAllText(@"ds\files_failed.txt", x + "\r\n");
                    }
                    else
                    {
                        Console.WriteLine("Compiled", x);
                        File.AppendAllText(@"ds\files_compiled.txt", Path.GetFileNameWithoutExtension(x) + "\t" + x + "\r\n");
                        foreach (var key in t.Keys)
                        {
                            var lst = t[key];
                            File.AppendAllLines(@"ds\" + key + "-ds.csv", lst);
                        }
                        Console.WriteLine(cntr++);
                    }
                });
            }
            else
            {
                db.ExecSql("delete from hpv_list;");
                directory.ToList().ForEach(x =>
                {
                    Console.WriteLine(filecounter++);
                    if (ignoredFIles.Contains(x))
                    {
                        File.AppendAllText(@"ds\registers\files_ingnored.txt", x + "\r\n");
                        return;
                    }

                    //if (cntr >= 30) return;
                    ////sqlconn = new SqlConnection(@"data source = '.\sqldev'; integrated seccurity=true;initial catalog=jhpiegodb_hpv");
                    ////sqlconn.Open();

                    var excel = new GetExcelAsDataTable3()
                    {
                        excelApp = excelapp,
                        addHeaderRow = cntr == 0,
                        fileName = x,
                        sqlConn = new SqlConnection(@"data source = '.\sqldev'; integrated security=true; initial catalog='jhpiegodb_hpv'")
                    };
                    var t = excel.Execute();
                    if(t==null && excel.fileSkipped)
                    {
                        Console.WriteLine("Skipped ${0}", x);
                    }
                    else if (t == null && excel.headersMatched)
                    {
                        Console.WriteLine("Matched", x);
                    }
                    else if (t == null)
                    {
                        Console.WriteLine("Not compiled", x);
                        File.AppendAllText(@"ds\registers\files_failed.txt", x + "\r\n");
                    }
                    else
                    {
                        Console.WriteLine("Compiled", x);
                        File.AppendAllText(@"ds\registers\files_compiled.txt", Path.GetFileNameWithoutExtension(x) + "\t" + x + "\r\n");
                        foreach (var key in t.Keys)
                        {
                            var lst = t[key];
                            File.AppendAllLines(@"ds\registers\hpv_register.csv", lst);
                        }
                        Console.WriteLine(cntr++);
                    }
                });
            }

            Console.WriteLine("Any key to exit");
            Console.ReadLine();
        }
    }
}
