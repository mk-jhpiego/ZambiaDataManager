using db_resyncer.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace db_resyncer
{
    class Program
    {
        static void Main(string[] args)
        {
            //we get the list of tables to sync
            var tableProcessors = TableHelper.getTableProcessors();

            //we read the data from the server 
            foreach(var table in tableProcessors)
            {
                var sql = 
                //var ds = table.getDbData();
            }
//float
//int
//nvarchar
//smallint
            //we spit data from the server to bits

            //upload the bits to some shared drive / cloud

            //we connect to the server

            //we connect to the localdb

            //we get the list of tables for syncing

            //we read data from the server to local

            //we push data from client to the server
        }
    }
}
