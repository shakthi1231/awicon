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
using Database;
using System.Web;
using System.Threading.Tasks;

namespace awicon.Controllers
{
    public class TemplatesController : ApiController
    {
        // GET: api/Templates
        [Route("api/column-names/{table}")]
        public HttpResponseMessage GetCategoryId(string table) {
            List<object> columnNames = new List<object>();
            var sqlConnection = SqlLib.getConnection();
            var commandText = "select COLUMN_NAME from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME = @table  AND TABLE_SCHEMA='dbo'";
            SqlCommand command = new SqlCommand(commandText, sqlConnection);
            command.Parameters.AddWithValue("@table", table);
            try
            {
                sqlConnection.Open();
                using (var sqlQueryResult = command.ExecuteReader())
                    if (sqlQueryResult != null)
                    {
                        Debug.WriteLine("SqlResult is successfull");


                        while (sqlQueryResult.Read())
                        {
                            var tableName = sqlQueryResult.GetValue(0).ToString();
                            columnNames.Add(tableName);

                        }

                        Debug.WriteLine("Printing all the tableNames...");
                        Debug.WriteLine(columnNames);
                        Debug.WriteLine("finished");
                    }

            }
            catch (Exception ex)
            {
                Debug.WriteLine("Ëxception is thrown");
                Debug.WriteLine(ex.Message);
            }
            var response = Request.CreateResponse(HttpStatusCode.OK);
            response.Content = new ObjectContent<List<object>>(columnNames, Configuration.Formatters.JsonFormatter, "application/json");
            return response;
        }


        public HttpResponseMessage Get(HttpRequestMessage request)
        {
           // var tableNames = new string[];
            var connection = @"Server=CWDNX52\SQLEXPRESS;Database=Reports;Trusted_Connection=True;";
            List<object> tableNames = new List<object>();
            using (var sqlConnection = new SqlConnection(connection))
            {
                var commandText = @"SELECT table_name FROM information_schema.tables";
                SqlCommand command = new SqlCommand(commandText, sqlConnection);
       
                try
                {
                    sqlConnection.Open();
                    using (var sqlQueryResult = command.ExecuteReader())
                        if (sqlQueryResult != null)
                        {
                            Debug.WriteLine("SqlResult is successfull");
   
                          
                            while ( sqlQueryResult.Read())
                            {   var tableName = sqlQueryResult.GetValue(0).ToString();
                                tableNames.Add(tableName);
                               
                            }

                            Debug.WriteLine("Printing all the tableNames...", tableNames);
                            Debug.WriteLine("finished");
                        }

                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Ëxception is thrown");
                    Debug.WriteLine(ex.Message);
                }
            }
            var response = Request.CreateResponse(HttpStatusCode.OK);
            response.Content = new ObjectContent<List<object>>(tableNames, Configuration.Formatters.JsonFormatter, "application/json");
            return response;

        }

        // GET: api/Templates/5
        public string Get(int id)
        {
            return "value";
        }

        // POST: api/Templates
        public HttpResponseMessage Post()
        {

           // var files = HttpContext.Current.Request.Params["files"];
            var table = HttpContext.Current.Request.Params["table"];
            var name = HttpContext.Current.Request.Params["name"];
            var map = HttpContext.Current.Request.Params["map"];
            var startRow = HttpContext.Current.Request.Params["startRow"];
            var httpRequest = HttpContext.Current.Request;
            Debug.WriteLine("Posting Data .....");
            Debug.WriteLine(httpRequest);
            if (httpRequest.Files.Count < 1)
            {
                Debug.WriteLine("Bad Request");
                return Request.CreateResponse(HttpStatusCode.BadRequest);
            }

            var postedFile = httpRequest.Files[0];
            byte[] file;
            var stream = postedFile.InputStream;
            using (var reader = new BinaryReader(stream))
            {
                file = reader.ReadBytes((int)stream.Length);
            }

            // TODO - Same FileData can be used to  store data in Database also
            // var filePath = HttpContext.Current.Server.MapPath("~/" + postedFile.FileName);
            //  postedFile.SaveAs(filePath);
            //Do whatever you want with filename and its binaray data.
            var connection = @"Server=CWDNX52\SQLEXPRESS;Database=Reports;Trusted_Connection=True;";
            using (var sqlConnection = new SqlConnection(connection))
            {
                var commandText = "insert into dbo.Reports([template],[map],[name],[table],[startRow]) values(@file , @map , @name , @table, @startRow)";

                SqlCommand command = new SqlCommand(commandText, sqlConnection);
                command.Parameters.Add("@file", SqlDbType.VarBinary, file.Length).Value = file;
                command.Parameters.AddWithValue("@map", map);
                command.Parameters.AddWithValue("@name", name);
                command.Parameters.AddWithValue("@table", table);
                command.Parameters.AddWithValue("@startRow", Int32.Parse(startRow));
                // Use AddWithValue to assign Demographics.
                // SQL Server will implicitly convert strings into XML.

                try
                {
                    sqlConnection.Open();
                    Int32 rowsAffected = command.ExecuteNonQuery();
                    Console.WriteLine("RowsAffected: {0}", rowsAffected);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            Debug.WriteLine(table);
            Debug.WriteLine("Post request for creating a template");
            Debug.WriteLine(name);
            Debug.WriteLine(map);


            Debug.WriteLine("Printed Template");
            return Request.CreateResponse(HttpStatusCode.Created);
        }

        // PUT: api/Templates/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE: api/Templates/5
        public void Delete(int id)
        {
        }
    }
}
