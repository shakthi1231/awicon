using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using Database;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Data;

using Newtonsoft.Json.Linq;
using ClosedXML.Excel;
using System.IO;
using Spire.Xls;
using System.Net.Http.Headers;

namespace awicon.Controllers
{
    public class ReportsController : ApiController
    {
        // GET: api/Reports
        public static void buildLayout(SqlDataReader file)
        {
              try { 
                           // var blob = new Byte[(file.GetBytes(0, 0, null, 0, int.MaxValue)];
                           // sqlQueryResult.GetBytes(0, 0, blob, 0, blob.Length);
                          //  using (var fs = new FileStream(varPathToNewLocation, FileMode.Create, FileAccess.Write))
                          //      fs.Write(blob, 0, blob.Length);
                            Debug.WriteLine("Completed.....");
                        

                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Ëxception is thrown");
                    Debug.WriteLine(ex.Message);
                }
            }

        private HttpResponseMessage writeToExcel(DataTable dt)
        {
            XLWorkbook Workbook = new XLWorkbook(@"c:\temp\Sample.xlsx");
            HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
            IXLWorksheet Worksheet = Workbook.Worksheets.First();
            Debug.WriteLine(Worksheet);
            IXLCell cellfornewdata = Worksheet.Cell(8, 1);
            cellfornewdata.InsertTable(dt.AsEnumerable());
            Worksheet.Row(8).Delete();
            Workbook.SaveAs(@"c:\temp\Sample.xlsx");

            try
            {
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(@"c:\temp\Sample.xlsx");
                workbook.ConverterSetting.SheetFitToPage = true;
                workbook.CustomDocumentProperties.Add("_MarkAsFinal", true);
                workbook.CustomDocumentProperties.Add("The Editor", "E-iceblue");
                workbook.CustomDocumentProperties.Add("Phone number1", 81705109);
                workbook.CustomDocumentProperties.Add("Revision number", 7.12);
                workbook.CustomDocumentProperties.Add("Revision date", DateTime.Now);
                workbook.SaveToFile(@"c:\temp\Sample.pdf", FileFormat.PDF);
                Process.Start(@"c:\temp\Sample.pdf");
                var pdf = @"c:\temp\Sample.pdf";
               
                var stream = new FileStream(pdf, FileMode.Open, FileAccess.Read);
                result.Content = new StreamContent(stream);
                result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
                return result;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Ëxception is thrown");
                Debug.WriteLine(ex.Message);
                return result;
            }
        }

        private DataTable getColumns(string tableName, string colArray)
        {
            DataTable dt = new DataTable();
            List<string> colNames = new List<string>();
            colArray =  colArray.Replace("[", "");
            colArray = colArray.Replace("]", "");

            Debug.WriteLine(colArray);
            string[] newColArray = colArray.Split(',');
            Debug.WriteLine(newColArray);

            for(int i=0; i < newColArray.Length; i++)
            {
                var column = JObject.Parse(newColArray[i]);
                Debug.WriteLine(column["name"].ToString());
                colNames.Add(column["name"].ToString());
            }

            string  colQuery = string.Join(",", colNames.ToArray());
            Debug.WriteLine(colNames);
            Debug.WriteLine("Printed colNames.......");
            Debug.WriteLine(colQuery);
            

            var connection = SqlLib.getConnection();
            var commandText = @"SELECT " + colQuery + " FROM [dbo].[" + tableName + "]";
            SqlCommand command = new SqlCommand(commandText, connection);

            command.Parameters.AddWithValue("@tableName", tableName);
            command.Parameters.AddWithValue("@colQuery", colQuery);

            try
            {
                connection.Open();
                using (var sqlQueryResult = command.ExecuteReader())
                    if (sqlQueryResult != null)
                    {

                        Debug.WriteLine("SqlResult is successfulr");
                        for (int i = 0; i < colNames.Count; i++)
                        {  
                            dt.Columns.Add(colNames[i]);
                        }
                        int count = 0;
                        while (sqlQueryResult.Read() && count < 300){
                            DataRow dataRow = dt.NewRow();
                            for (int i = 0; i < colNames.Count; i++)
                            {
                                dataRow[colNames[i]] = sqlQueryResult[colNames[i]].ToString();
                            }
                            count++;
                            dt.Rows.Add(dataRow);

                        }
                        //var reportObj = sqlQueryResult.Read();

                        //var template = sqlQueryResult["template"];
                        //var startRow = Int32.Parse(sqlQueryResult["startRow"].ToString());
                        //var table = sqlQueryResult["table"].ToString();
                        //var map = sqlQueryResult["map"].ToString();


                       // DataTable dt = getColumns(table, map);

                    }

            }
            catch (Exception ex)
            {
                Debug.WriteLine("Ëxception is thrown");
                Debug.WriteLine(ex.Message);
            }
            foreach (DataRow dataRow in dt.Rows)
            {
                foreach (var item in dataRow.ItemArray)
                {
                    Debug.WriteLine(item);
                }
            }
            return dt;
        }


        public HttpResponseMessage Get(string report)
        {
            
            var connection = SqlLib.getConnection();
            var commandText = @"SELECT * FROM [dbo].[Reports] WHERE [name] = @name";
            HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
            SqlCommand command = new SqlCommand(commandText, connection);
            command.Parameters.AddWithValue("@name", report);
            try
            {
                connection.Open();
                using (var sqlQueryResult = command.ExecuteReader())
                    if (sqlQueryResult != null)
                    {
                        Debug.WriteLine("SqlResult is successfulr");
                        var reportObj =  sqlQueryResult.Read();
                        byte[] template = (byte[])sqlQueryResult["template"];
                        var startRow = Int32.Parse(sqlQueryResult["startRow"].ToString());
                        var table = sqlQueryResult["table"].ToString();
                        var map = sqlQueryResult["map"].ToString();


                        using (var fs = new FileStream(@"c:\temp\Sample.xlsx", FileMode.Create, FileAccess.Write))
                            fs.Write(template, 0, template.Length);
                        DataTable dt = getColumns(table,map);
                          result = writeToExcel(dt);

                    }

            }
            catch (Exception ex)
            {
                Debug.WriteLine("Ëxception is thrown");
                Debug.WriteLine(ex.Message);
            }
            return result;
        }

        // GET: api/Reports/5
        public string Get(int id)
        {
            return "value";
        }

        // POST: api/Reports
        public void Post([FromBody]string value)
        {
        }

        // PUT: api/Reports/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE: api/Reports/5
        public void Delete(int id)
        {
        }
    }
}
