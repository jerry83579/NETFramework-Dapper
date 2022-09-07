using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using OpenDocumentLib.sheet;
using System.Drawing;
using System.Web.Http;
using System.IO;
using System.Net.Http.Headers;
using FileTest.Models;
using FileTest.Helpers;
using GdalReadSHP;
using OSGeo.OGR;
using OSGeo.OSR;
using Microsoft.SharePoint.Client;
using OSGeo.GDAL;
using FieldType = OSGeo.OGR.FieldType;
using Driver = OSGeo.OGR.Driver;
using Feature = OSGeo.OGR.Feature;
using System.Runtime.InteropServices;
using Microsoft.SqlServer.Types;
using System.Data.SqlTypes;

namespace FileTest.Controllers
{
    /// <summary>
    /// 
    /// </summary>
    public class OdsController : ApiController
    {
        public SqlConn sqlConn;
        public ShpRead m_Shp;
        public static Layer pLayer = null;
        //public SqlConnection myConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString);

        public OdsController()
        {
         sqlConn = new SqlConn();
         m_Shp = new ShpRead();
        }

        //ods
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

        //shp
        [DllImport("gdal204.dll", EntryPoint = "OGR_F_GetFieldAsString", CallingConvention = CallingConvention.Cdecl)]
        public extern static System.IntPtr OGR_F_GetFieldAsString(HandleRef handle, int index);

        /// <summary>
        /// 創建要素圖層
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="geometryType"></param>
        /// <param name="spatialReference"></param>
        /// <returns></returns>
        public  static Layer CreateLayer(string filePath, wkbGeometryType geometryType, SpatialReference spatialReference)
        {
            // 數據源
            Driver pDriver = Ogr.GetDriverByName("ESRI Shapefile");
            DataSource pDataSource = pDriver.CreateDataSource(filePath, null);

            // 創建圖層
            if (geometryType == wkbGeometryType.wkbPoint)
            {
                pLayer = pDataSource.CreateLayer(Path.GetFileNameWithoutExtension(filePath), spatialReference, wkbGeometryType.wkbPoint, null);
            }
            else if (geometryType == wkbGeometryType.wkbLineString)
            {
                pLayer = pDataSource.CreateLayer(Path.GetFileNameWithoutExtension(filePath), spatialReference, wkbGeometryType.wkbLineString, null);
            }
            else
            {
                pLayer = pDataSource.CreateLayer(Path.GetFileNameWithoutExtension(filePath), spatialReference, wkbGeometryType.wkbPolygon, null);
            }
            return pLayer;
        }

        //========================================================================= Import & Export =========================================================================
        /// <summary>
        /// 讀取資料庫寫入 Shp
        /// </summary>
        [Route("get/exportShp")]
        public void GetExportShp()
        {
            string filePath = @"D:\shapefile";
            string newPath = @"D:\shapefile\shapefile.shp";
            m_Shp.InitinalGdal();
            //創建圖層
            CreateLayer(filePath, wkbGeometryType.wkbPoint, null);

            //獲取數據源
            Driver pDriver = Ogr.GetDriverByName("ESRI Shapefile");
            DataSource pDataSource = pDriver.Open(newPath, 1);

            //創建欄位
            FieldDefn pFieldDefn = new FieldDefn("X", FieldType.OFTString);
            pFieldDefn.SetWidth(200);
            pLayer.CreateField(pFieldDefn, 1);

            pFieldDefn = new OSGeo.OGR.FieldDefn("Y", FieldType.OFTString);
            pFieldDefn.SetWidth(200);
            pLayer.CreateField(pFieldDefn, 1);

            pFieldDefn = new OSGeo.OGR.FieldDefn("C", FieldType.OFTString);
            pFieldDefn.SetWidth(200);
            pLayer.CreateField(pFieldDefn, 1);

            //獲得欄位索引值來新增資料
            OSGeo.OGR.FeatureDefn pFeatureDefn = pLayer.GetLayerDefn();
            int fieldIndex_A = pFeatureDefn.GetFieldIndex("X");
            int fieldIndex_B = pFeatureDefn.GetFieldIndex("Y");
            int fieldIndex_C = pFeatureDefn.GetFieldIndex("Z");

            // 插入要素
            OSGeo.OGR.Feature pFeature = new Feature(pFeatureDefn);
            OSGeo.OGR.Geometry geometry = Geometry.CreateFromWkt("POINT(203537 2636875)");
            pFeature.SetGeometry(geometry);
            pFeature.SetField(fieldIndex_A, "11");
            pFeature.SetField(fieldIndex_B, "1");
            pFeature.SetField(fieldIndex_C, "张三");
            pLayer.CreateFeature(pFeature);

            //geometry = Geometry.CreateFromWkt("POINT(203537 2536875)");
            //pFeature.SetGeometry(geometry);
            //pFeature.SetField(fieldIndex_A, "12");
            //pFeature.SetField(fieldIndex_B, "2");
            //pFeature.SetField(fieldIndex_C, "张四");
            //pLayer.CreateFeature(pFeature);
            //i++;
            ////geometry = Geometry.CreateFromWkt("POINT(11 110)");
            ////pFeature.SetGeometry(geometry);
            ////pFeature.SetField(fieldIndex_A, 1);
            ////pFeature.SetField(fieldIndex_B, "211032");
            ////pFeature.SetField(fieldIndex_C, "李四");
            ////pLayer.CreateFeature(pFeature);
        }


        /// <summary>
        /// 獲取 Shp 資料寫入資料庫
        /// </summary>
        [Route("get/importShp")]
        public void GetImportShp()
        {
            string filePath = @"D:\shp\shop.shp";
            m_Shp.InitinalGdal();
            // 数据源
            Driver pDriver = Ogr.GetDriverByName("ESRI Shapefile");
            DataSource pDataSource = pDriver.Open(filePath, 0);
            Layer pLayer = pDataSource.GetLayerByName(Path.GetFileNameWithoutExtension(filePath));
            pLayer.GetSpatialRef();
            // 字段结构
            Feature pFeature = pLayer.GetNextFeature();
            FeatureDefn pFeatureDefn = pLayer.GetLayerDefn();
            int fieldsCount = pFeature.GetFieldCount();

            List<Dictionary<string, object>> listDic = new List<Dictionary<string, object>>();
            // 遍历要素
            while (pFeature != null)
            {
                Dictionary<string, object> dicData = new Dictionary<string, object>();
                pFeature.GetGeometryRef().ExportToWkt(out string spatialReferenceString);
                dicData.Add("Geo", spatialReferenceString);
                for (int i = 0; i < fieldsCount; i++)
                {
                    FieldDefn pFieldDefn = pFeature.GetFieldDefnRef(i);
                    string fieldName = pFieldDefn.GetName();
                    switch (pFieldDefn.GetFieldType())
                    {
                        //case FieldType.OFTString:
                        //    {
                        //        Console.WriteLine(pFeature.GetFieldAsString(i));
                        //        dicData.Add(fieldName, pFeature.GetFieldAsInteger(i));
                        //    }
                        //    break;
                        case FieldType.OFTInteger:
                            {
                                Console.WriteLine(pFeature.GetFieldAsInteger(i));
                                dicData.Add(fieldName, pFeature.GetFieldAsInteger(i));
                            }
                            break;
                        case FieldType.OFTInteger64:
                            {
                                Console.WriteLine(pFeature.GetFieldAsInteger64(i));
                                dicData.Add(fieldName, pFeature.GetFieldAsInteger64(i));
                            }
                            break;
                        case FieldType.OFTReal:
                            {
                                Console.WriteLine(pFeature.GetFieldAsDouble(i));
                                dicData.Add(fieldName, pFeature.GetFieldAsDouble(i));
                            }
                            break;
                        case FieldType.OFTString:
                            {
                                //string fieldName = pFieldDefn.GetName();
                                int fieldIndex = pFeatureDefn.GetFieldIndex(fieldName);
                                IntPtr pIntPtr = OGR_F_GetFieldAsString(OSGeo.OGR.Feature.getCPtr(pFeature), fieldIndex);
                                Console.WriteLine(Marshal.PtrToStringAnsi(pIntPtr));
                                dicData.Add(fieldName, pFeature.GetFieldAsString(i));
                                //Marshal.PtrToStringAnsi(pIntPtr)
                            }
                            break;
                    }
                }
                Console.WriteLine("----------------------------------------------");
                listDic.Add(dicData);
                pFeature = pLayer.GetNextFeature();
            }


            pDriver.Register();
            pDataSource.FlushCache();
            pDataSource.Dispose();
            //pDriver.DeleteDataSource(filePath);
            pDriver.Dispose();
           
            string conStr = @"INSERT INTO Info (Name, Food, Address, Phone, Location) VALUES(@Name, @Food, @Address, @Phone, @Location)";
            var data = new List<Info>();
            int listDicLength = listDic.Count;
            for (int i = 0; i< listDicLength; i++)
            {
                data.Add(new Info()
                {
                    Name = listDic[i]["Name"].ToString(),
                    Food = listDic[i]["Food"].ToString(),
                    Address = listDic[i]["Address"].ToString(),
                    Phone = listDic[i]["Phone"].ToString(),
                    Location = SqlGeometry.STGeomFromText(new SqlChars(listDic[i]["Geo"].ToString()), 4326)
                }); 
            }
            sqlConn.DbExecute(conStr, data);
        }

        /// <summary>
        /// ods
        /// </summary>
        [Route("get/import")]
            public void GetImport()
            {
            Calc.LogPath = @"D:\log";
            var path = @"D:\New.ods";
            var row = 1;
            using (var calc = new Calc(path))
            {
            var workSheet = calc.Tables[1];
            int columnsLength = workSheet.ColumnCount-2;
            int rowsLength = workSheet.RowCount;
            string conStr = @"INSERT INTO Info (Name, Food, Address, Phone, Lat, Longitude, Location) VALUES(@Name, @Food, @Address, @Phone, @Lat, @Longitude, @Location)";
            var data = new List<Info>();
                
                for (int i = 0; i < columnsLength; i++)
                {
                    double Lat = workSheet[row, 5].Formula.ToDouble();
                    double Longitude = workSheet[row, 6].Formula.ToDouble();
                    data.Add(new Info()
                    {
                        Name = workSheet[row, 1].Formula,
                        Food = workSheet[row, 2].Formula,
                        Address = workSheet[row, 3].Formula,
                        Phone = workSheet[row, 4].Formula,
                        Lat = Lat,
                        Longitude = Longitude,
                        Location = SqlGeometry.Point(Lat, Longitude, 4326)
                });
                    row++;
                }
            int result = sqlConn.DbExecute(conStr, data);
            };
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
                foreach (var rowData in dataset)
                {
                    foreach (var field in rowData)
                    {
                        if (row <= header)
                        {
                        tb[row, column].Formula = field.Key;
                        }
                        tb[row + 1, column].Formula = Convert.ToString(field.Value);
                        column ++;
                    }
                    row ++;
                    column = 0;
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
