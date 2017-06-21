using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Data;
using System.Net.Http;
using System.Web.Http;
using ClosedXML.Excel;
using System.Diagnostics;

namespace awicon.Controllers
{
    public class ExcelController : ApiController
    {
        // GET: api/Excel
        public void Get()
        {
            XLWorkbook Workbook = new XLWorkbook(@"c:\temp\Sample.xlsx");
            IXLWorksheet Worksheet = Workbook.Worksheets.First();
            Debug.WriteLine(Worksheet);
            int NumberOfLastRow = Worksheet.LastRowUsed().RowNumber();
            IXLCell CellForNewData = Worksheet.Cell(8, 1);
            DataTable datatable = new DataTable();
            datatable.Columns.Add("Name");
            datatable.Columns.Add("Marks");
            datatable.Columns.Add("Rank");
            var i = 10;
            while(i > 0)
            {
                DataRow dr = datatable.NewRow();
                dr["Name"] = "shakthi";
                dr["Marks"] = "shiva";
                dr["Rank"] = "sharan";
     
                datatable.Rows.Add(dr);
                --i;
            }
            CellForNewData.InsertTable(datatable.AsEnumerable());
            Worksheet.Row(8).Delete();
            Workbook.Save();

        }

        // GET: api/Excel/5
        public string Get(int id)
        {
            return "value";
        }

        // POST: api/Excel
        public void Post([FromBody]string value)
        {
        }

        // PUT: api/Excel/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE: api/Excel/5
        public void Delete(int id)
        {
        }
    }
}
