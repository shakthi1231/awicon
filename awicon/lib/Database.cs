using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using ClosedXML.Excel;
using System.IO;
using System.Data;
using Spire.Xls;
using System.Data.SqlClient;
using System.Diagnostics;

namespace Database
{
    public class SqlLib
    {

        public static SqlConnection getConnection()
        {
            var connection = @"Server=CWDNX52\SQLEXPRESS;Database=Reports;Trusted_Connection=True;";
            var sqlConnection = new SqlConnection(connection);
           
                return sqlConnection;
            

        }
    }
}