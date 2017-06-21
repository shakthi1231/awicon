using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;
using ClosedXML.Excel;
using System.IO;
using Spire.Xls;
using System.Data.SqlClient;
using System.Diagnostics;


namespace awicon.Controllers
{
    //[Authorize]
    public class ValuesController : ApiController
    {
        public static void databaseFileRead(string varPathToNewLocation)
        {
            Debug.WriteLine("DatabaseFileRead.........");
            var connection = @"Server=CWDNX52\SQLEXPRESS;Database=Reports;Trusted_Connection=True;";
            using (var sqlConnection = new SqlConnection(connection))
            {
                var startRow = 8;
               var commandText = @"SELECT [template] FROM [dbo].[Reports] WHERE [startRow] = @startRow";
         
                SqlCommand command = new SqlCommand(commandText, sqlConnection);
               command.Parameters.AddWithValue("@startRow", startRow);
                // Use AddWithValue to assign Demographics.
                // SQL Server will implicitly convert strings into XML.

                try
                {
                    sqlConnection.Open();
                    using (var sqlQueryResult = command.ExecuteReader())
                        if (sqlQueryResult != null)
                        {
                            Debug.WriteLine("SqlResult is successfulr");
                            sqlQueryResult.Read();
                            var blob = new Byte[(sqlQueryResult.GetBytes(0, 0, null, 0, int.MaxValue))];
                            sqlQueryResult.GetBytes(0, 0, blob, 0, blob.Length);
                            using (var fs = new FileStream(varPathToNewLocation, FileMode.Create, FileAccess.Write))
                                fs.Write(blob, 0, blob.Length);
                            Debug.WriteLine("Completed.....");
                        }
                    
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Ëxception is thrown");
                    Debug.WriteLine(ex.Message);
                }
            }
        }

        public static void databaseFilePut(string varFilePath)
        {
            byte[] file;
            using (var stream = new FileStream(varFilePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = new BinaryReader(stream))
                {
                    file = reader.ReadBytes((int)stream.Length);
                }
            }

            var connection = @"Server=CWDNX52\SQLEXPRESS;Database=Reports;Trusted_Connection=True;";
            using (var sqlConnection = new SqlConnection(connection))
            {
                var commandText = "insert into dbo.Wyibe([file],[Id]) values(@file, @Id)";
              
                SqlCommand command = new SqlCommand(commandText, sqlConnection);
                command.Parameters.Add("@File", System.Data.SqlDbType.VarBinary, file.Length).Value = file;
                command.Parameters.AddWithValue("@Id", 1);
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
        }

        public void Get()
        {
            //return new string[] { "value1", "value2" };
            //string path = Directory.GetCurrentDirectory();
            //Console.WriteLine("------------------------------------", path);
            //string target = @"c:\temp";
            //Console.WriteLine("The current directory is {0}", path);
            //if (!Directory.Exists(target))
            //{
            //    Directory.CreateDirectory(target);
            //}

            //// Change the current directory.
            //Environment.CurrentDirectory = (target);
            //if (path.Equals(Directory.GetCurrentDirectory()))
            //{
            //    Console.WriteLine("You are in the temp directory.");
            //}
            //else
            //{
            //    Console.WriteLine("You are not in the temp directory.");
            //}
            //var workbook = new XLWorkbook();
            //string fileName = "HelloWorld.xlsx";
            //var worksheet = workbook.Worksheets.Add("Sample Sheet");
            //worksheet.Cell("A1").Value = "Hello World!";
            //workbook.SaveAs("HelloWorld.xlsx");

            //Workbook workbook1 = new Workbook();
            //workbook1.LoadFromFile(@"c:\temp\HelloWorld.xlsx");
            //workbook1.ConverterSetting.SheetFitToPage = true;
            //workbook1.SaveToFile(@"c:\temp\HelloWorld.pdf", FileFormat.PDF);
            //System.Diagnostics.Process.Start(@"c:\temp\HelloWorld.pdf");
            //var pdf = @"c:\temp\HelloWorld.pdf";
            //HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
            //var stream = new FileStream(pdf, FileMode.Open, FileAccess.Read);
            //result.Content = new StreamContent(stream);
            //result.Content.Headers.ContentType =
            //    new MediaTypeHeaderValue("application/pdf");
            //return result;

          //  databaseFilePut(@"c:\temp\HelloWorld.xlsx");
            databaseFileRead(@"c:\temp\first.xlsx");
            return;


        }

        // GET api/values/5
        public string Get(int id)
        {
            return "value";
        }

        // POST api/values
        public void Post([FromBody]string value)
        {
        }

        // PUT api/values/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/values/5
        public void Delete(int id)
        {
        }
    }
}
