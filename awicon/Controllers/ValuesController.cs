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


namespace awicon.Controllers
{
    //[Authorize]
    public class ValuesController : ApiController
    {
       
        public HttpResponseMessage Get()
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

            Workbook workbook1 = new Workbook();
            workbook1.LoadFromFile(@"c:\temp\HelloWorld.xlsx");
            workbook1.ConverterSetting.SheetFitToPage = true;
            workbook1.SaveToFile(@"c:\temp\HelloWorld.pdf", FileFormat.PDF);
            System.Diagnostics.Process.Start(@"c:\temp\HelloWorld.pdf");
            var pdf = @"c:\temp\HelloWorld.pdf";
            HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
            var stream = new FileStream(pdf, FileMode.Open, FileAccess.Read);
            result.Content = new StreamContent(stream);
            result.Content.Headers.ContentType =
                new MediaTypeHeaderValue("application/pdf");
            return result;
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
