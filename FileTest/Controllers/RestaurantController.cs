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
        public SqlConn sqlConn;
        public SqlConnection myConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString);

        public OdsController()
        {
         sqlConn = new SqlConn();
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

    

            [Route("get/import")]
            public void GetImport()
            {
            Calc.LogPath = @"D:\log";
            var path = @"D:\test.ods";
            using (var calc = new Calc(path))
            {
            string conStr = @"INSERT INTO Info (Name, Food) VALUES(123, 123)";
            int result = myConnection.Execute(conStr);
            }
            
        
            }

        /// <summary>
        /// 資料庫資料匯出到 Ods 檔案
        /// </summary>
        /// <returns></returns>
        [Route("get/export")]
        public HttpResponseMessage GetExport()
        {
            Calc.LogPath = @"D:\log";
            var outputPath = @"D:\New.ods";
            using (var calc = new Calc())
            {
                var tb = calc.Tables.AddNew("新的工作表");
                // 樣式
                Font headerFont = new Font("標楷體", 28, FontStyle.Bold),
                colFont = new Font("標楷體", 14, FontStyle.Bold),
                rowFont = new Font("標楷體", 12);

                var line = new Line() { Color = Color.Black, OuterWidth = 20 };
                var conStr =
                @"
                SELECT * 
                FROM dbo.Info 
                ";

                int header = 0;
                int row = 0;
                int column = 0;
                var dataset = sqlConn.DbQuery(conStr);
                foreach (var columnData in dataset)
                {
                    foreach (var field in columnData)
                    {
                        if (column <= header)
                        {
                        tb[column, row].Formula = field.Key;
                        }
                        tb[column+1, row].Formula = Convert.ToString(field.Value);
                        row ++;
                    }
                    column ++;
                    row = 0;
                }

                calc.SaveAs(outputPath);
                calc.Close();
            }
            return FileResult(outputPath, "application/vnd.oasis.opendocument.spreadsheet");
        }


        /// <summary>
        /// 修改本地 Ods 檔案
        /// </summary>
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
            }
        }

        /// <summary>
        /// 取得資料庫資料
        /// </summary>
        /// <returns></returns>
        [Route("get/database")]
        public IEnumerable<dynamic> GetDataBase()
        {
            var conStr =
            @"
            SELECT * 
            FROM dbo.Info 
            ";
            var data = sqlConn.DbQuery(conStr);
            return data;
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
