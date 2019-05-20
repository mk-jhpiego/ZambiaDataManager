using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HPVExcelReader
{
    enum operation{
        Microplans,
        Registers,
        Vaccines
    }
    class Program
    {
        static void Main(string[] args)
        {
            //var getValues = true;
            Console.WriteLine("Press any key to start");
            var g = Console.ReadLine();
            var registers_toIgnore = File.ReadAllLines("toIgnore_registers.txt");

            var microplans_toCompile = File.ReadAllLines("microplans-toCompile-20190519.txt");
            var microplans_toIgnore =  File.ReadAllLines("microplans-toIgnore-20190519.txt");            
            File.AppendAllText("headermatches.csv", "header,rowid,columnidd,districtname,filename\r\n");

            Microsoft.Office.Interop.Excel.Application excelapp = new Microsoft.Office.Interop.Excel.Application() { Visible = false };
            
            var folder = @"K:\development\github\ZambiaDataManager\HPVExcelReader\HPVExcelReader\bin\hpvdocs";
            var directory = Directory.GetFiles(folder, "*.*", SearchOption.AllDirectories);
            //Console.WriteLine(string.Format("Files found: {0}", directory.Length));
            var cntr = 0;
            var filecounter = 0;
            //verified-microplans 20190519.txt
            var processingMicroplans = operation.Microplans;
            if (processingMicroplans== operation.Microplans|| processingMicroplans == operation.Vaccines)
            {
                if (processingMicroplans == operation.Vaccines)
                {
                    db.ExecSql("delete from vaccines_list;");
                }
                if (processingMicroplans == operation.Microplans)
                {
                    db.ExecSql("delete from girls_primary;delete from girls_secondary;delete from girls_community;");
                }
                folder = @"K:\development\github\ZambiaDataManager\HPVExcelReader\HPVExcelReader\bin\microplans-finalist";
                directory = Directory.GetFiles(folder, "*.*", SearchOption.AllDirectories);
                Console.WriteLine(string.Format("Files found: {0}", directory.Length));
                directory.ToList().ForEach(x =>
                {
                    Console.WriteLine(filecounter++);
                    if (microplans_toIgnore.Contains(x) || microplans_toCompile.Contains(x))
                    {
                        Path.GetFileName(x);

                        //var fileinfo = new FileInfo(x);
                        //var targetfolder = @"K:\development\github\ZambiaDataManager\HPVExcelReader\HPVExcelReader\bin\Debug\ds\microplans";
                        //fileinfo.CopyTo(targetfolder+"\\"+ Path.GetFileName(x));
                        if (microplans_toIgnore.Contains(x))
                        {
                            //File.AppendAllText(@"ds\files_ingnored.txt", x + "\r\n");
                            return;
                        }
                    }
                    else
                    {
                        File.AppendAllText(@"ds\files_ignored.txt", x + "\r\n");
                        return;
                    }
                    if (processingMicroplans == operation.Microplans)
                    {
                        //if (cntr >= 30) return;
                        var excel = new GetMicroplans()
                        {
                            excelApp = excelapp,
                            addHeaderRow = cntr == 0,
                            fileName = x,
                            sqlConn = new SqlConnection(@"data source = '.\sqldev'; integrated security=true; initial catalog='jhpiegodb_hpv'")
                        };
                        var t = excel.Execute();
                        if (t == null)
                        {
                            Console.WriteLine("Not compiled", x);
                            File.AppendAllText(@"ds\microplans\files_failed.txt", x + "\r\n");
                        }
                        else
                        {
                            Console.WriteLine("Compiled", x);
                            File.AppendAllText(@"ds\microplans\files_compiled.txt", Path.GetFileNameWithoutExtension(x) + "\t" + x + "\r\n");
                            Console.WriteLine("Uncomment this");
                            foreach (var key in t.Keys)
                            {
                                var lst = t[key];
                                File.AppendAllLines(@"ds\microplans\" + key + "-ds.csv", lst);
                            }
                            Console.WriteLine(cntr++);
                        }
                    }
                    else
                    {                        
                        var excel = new GetVaccinesHelper()
                        {
                            excelApp = excelapp,
                            addHeaderRow = cntr == 0,
                            fileName = x,
                            sqlConn = new SqlConnection(@"data source = '.\sqldev'; integrated security=true; initial catalog='jhpiegodb_hpv'")
                        };
                        var t = excel.Execute();
                        if (t == null && excel.fileSkipped)
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
                            File.AppendAllText(@"ds\vaccines\files_failed.txt", x + "\r\n");
                        }
                        else
                        {
                            Console.WriteLine("Compiled", x);
                            File.AppendAllText(@"ds\vaccines\files_compiled.txt", Path.GetFileNameWithoutExtension(x) + "\t" + x + "\r\n");
                            foreach (var key in t.Keys)
                            {
                                var lst = t[key];
                                File.AppendAllLines(@"ds\vaccines\vaccine_total.csv", lst);
                            }
                            Console.WriteLine(cntr++);
                        }
                    }
                });
            }
            else if (processingMicroplans == operation.Registers)
            {
                db.ExecSql("delete from hpv_list;");
                directory.ToList().ForEach(x =>
                {
                    Console.WriteLine(filecounter++);
                    if (registers_toIgnore.Contains(x))
                    {
                        File.AppendAllText(@"ds\registers\files_ingnored.txt", x + "\r\n");
                        return;
                    }

                    //if (cntr >= 30) return;
                    ////sqlconn = new SqlConnection(@"data source = '.\sqldev'; integrated seccurity=true;initial catalog=jhpiegodb_hpv");
                    ////sqlconn.Open();

                    var excel = new GetRegistersHelper()
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
