using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using OpenDocumentLib.sheet;
using System.Drawing;
using System.Web.Http;
using OfficeDevPnP.Core.Utilities;
using System.Threading.Tasks;
using System.IO;
using System.Net.Http.Headers;
using System.Web;
using Microsoft.Graph;
using System.Data.Common;
using System.Configuration;
using System.Data;
using FileTest.Models;
using System.Data.SqlClient;
using FileTest.Helpers;
using Dapper;

namespace FileTest.Controllers
{
    /// <summary>
    /// 
    /// </summary>
    public class OdsController : ApiController
    {
      
        public OdsController() { 
            //SqlConn sqlCon = new SqlConn();
    }
       


        public static HttpResponseMessage FileResult(string filePath, string mime = "application/octet-stream")
        {
            var result = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StreamContent(System.IO.File.OpenRead(filePath))
            };
            result.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment")
            {
                FileName = Path.GetFileName(filePath)
            };
            result.Content.Headers.ContentType = new MediaTypeHeaderValue(mime);
            return result;
        }


        //private IDbConnection _db = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString);


        // DOWNLOAD api/ods
        //[Route("get/database")]
        //public List<dynamic> GetDataBase()
        //{
        //    using (var tmpCon = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString))
        //    {
        //        tmpCon.Open();
        //        var StudentList = tmpCon.Query("Select * From dbo.Info").ToList();
        //        return StudentList;
        //    }
        //}

        //DOWNLOAD api/ods
        [Route("get/database")]
        public List<dynamic> GetDataBase()
        {
            var conStr =
            @"
            SELECT * 
            FROM dbo.Info 
            Where Id = 2
            ";
            var sqlCon = new SqlConn();
            var database = sqlCon.DbQuery(conStr);
            return database;
        }


        // DOWNLOAD api/ods
        [Route("get/download")]
        public HttpResponseMessage GetDownload()
        {
            var outputPath = @"D:\output.ods";
            return FileResult(outputPath, "application/vnd.oasis.opendocument.spreadsheet");
        }


        // GET api/ods
        [Route("get/ods")]
        public void Get()
        {
            Calc.LogPath = @"D:\log";
            var path = @"D:\test.ods";
            var outputPath = @"D:\output.ods";
            using (var calc = new Calc(path))
            {
                var tb = calc.Tables.AddNew("新的工作表");
                // 樣式
                Font headerFont = new Font("標楷體", 28, FontStyle.Bold),
                colFont = new Font("標楷體", 14, FontStyle.Bold),
                rowFont = new Font("標楷體", 12);
                var line = new Line() { Color = Color.Black, OuterWidth = 20 };
                int startRow = 0; // 要調整明細位置只要改這個數值就好了
                tb[startRow, 3].Formula = "12345";
                calc.SaveAs(outputPath);
                calc.Close();
                //return HttpHelper.FileResult(filePath, "application/vnd.oasis.opendocument.spreadsheet");
            }
        }

        // GET api/values/5
        public string Get(int id)
        {
            return "value";
        }

        // POST api/values
        public void Post([FromBody] string value)
        {
        }

        // PUT api/values/5
        public void Put(int id, [FromBody] string value)
        {
        }

        // DELETE api/values/5
        public void Delete(int id)
        {
        }
    }
}
