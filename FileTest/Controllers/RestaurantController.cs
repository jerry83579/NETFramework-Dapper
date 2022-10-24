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
using Feature = OSGeo.OGR.Feature;
using Driver = OSGeo.OGR.Driver;
using System.Runtime.InteropServices;
using Microsoft.SqlServer.Types;
using System.Data.SqlTypes;
using Microsoft.Graph;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Drive;
using unoidl.com.sun.star.sdbc;
using System.Reflection;
using NLog;
using System.Threading.Tasks;
using System.Web.Mvc;
using Microsoft.AspNetCore.Mvc;
using RouteAttribute = System.Web.Http.RouteAttribute;
using Microsoft.AspNetCore.Http;
using System.Web.Http.Cors;
using HttpPostAttribute = System.Web.Http.HttpPostAttribute;
using System.Web;
using HttpContext = System.Web.HttpContext;
using Newtonsoft.Json.Linq;
using Directory = System.IO.Directory;
using System.Text;
using System.Collections.ObjectModel;
using FormCollection = System.Web.Mvc.FormCollection;
using HttpGetAttribute = System.Web.Http.HttpGetAttribute;
using System.Xml;
using System.Xml.Linq;
using System.Linq;
using OfficeDevPnP.Core.Utilities;
using Microsoft.Win32;

namespace FileTest.Controllers
{
    [EnableCors(origins: "*", headers: "*", methods: "*", SupportsCredentials = true)]
    public class OdsController : ApiController
    {
        public SqlConn sqlConn;
        public ShpRead m_Shp;
        public static Driver pDriver;
        public static DataSource dataSource;

        //shp
        [DllImport("gdal204.dll", EntryPoint = "OGR_F_GetFieldAsString", CallingConvention = CallingConvention.Cdecl)]
        public static extern System.IntPtr OGR_F_GetFieldAsString(HandleRef handle, int index);

        public OdsController()
        {
            sqlConn = new SqlConn();
            m_Shp = new ShpRead();
        }

        #region 下載檔案

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

        #endregion 下載檔案

        #region 創建圖層

        /// <summary>
        /// 創建要素圖層
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="geometryType"></param>
        /// <param name="spatialReference"></param>
        /// <returns></returns>
        public static Layer CreateLayer(string direPath, wkbGeometryType geometryType, SpatialReference spatialReference)
        {
            Layer pLayer = null;
            string[] options = { "ENCODING=UTF-8" };
            // 數據源
            dataSource = pDriver.CreateDataSource(direPath, options);

            // 創建圖層
            if (geometryType == wkbGeometryType.wkbPoint)
            {
                pLayer = dataSource.CreateLayer(Path.GetFileNameWithoutExtension(direPath), spatialReference, wkbGeometryType.wkbPoint, new string[] { "ENCODING=UTF-8" });
            }
            else if (geometryType == wkbGeometryType.wkbLineString)
            {
                pLayer = dataSource.CreateLayer(Path.GetFileNameWithoutExtension(direPath), spatialReference, wkbGeometryType.wkbLineString, new string[] { "ENCODING=UTF-8" });
            }
            else
            {
                pLayer = dataSource.CreateLayer(Path.GetFileNameWithoutExtension(direPath), spatialReference, wkbGeometryType.wkbPolygon, new string[] { "ENCODING=UTF-8" });
            }
            return pLayer;
        }

        #endregion 創建圖層

        #region 新增點位

        public static void InsertFeatures(string filePath, string testPath)
        {
            // 数据源
            Driver pDriver = Ogr.GetDriverByName("ESRI Shapefile");
            DataSource pDataSource = pDriver.Open(filePath, 1);
            Layer pLayer = pDataSource.GetLayerByName(Path.GetFileNameWithoutExtension(testPath));

            //獲得欄位索引值來新增資料
            FeatureDefn pFeatureDefn = pLayer.GetLayerDefn();
            int fieldIndex_A = pFeatureDefn.GetFieldIndex("A");
            int fieldIndex_B = pFeatureDefn.GetFieldIndex("B");
            int fieldIndex_C = pFeatureDefn.GetFieldIndex("C");

            // 插入要素
            Feature pFeature = new Feature(pFeatureDefn);
            Geometry geometry = Geometry.CreateFromWkt("POINT(203537 2636875)");
            pFeature.SetGeometry(geometry);
            pFeature.SetField(fieldIndex_A, "1");
            pFeature.SetField(fieldIndex_B, "SADFFDS");
            pFeature.SetField(fieldIndex_C, "SADFSADF");
            pLayer.CreateFeature(pFeature);

            geometry = Geometry.CreateFromWkt("POINT(203537 2536875)");
            pFeature.SetGeometry(geometry);
            pFeature.SetField(fieldIndex_A, "2");
            pFeature.SetField(fieldIndex_B, "食物B");
            pFeature.SetField(fieldIndex_C, "北平路");
            pLayer.CreateFeature(pFeature);

            geometry = Geometry.CreateFromWkt("POINT(11 110)");
            pFeature.SetGeometry(geometry);
            pFeature.SetField(fieldIndex_A, "3");
            pFeature.SetField(fieldIndex_B, "食物C");
            pFeature.SetField(fieldIndex_C, "河南路");
            pLayer.CreateFeature(pFeature);
        }

        #endregion 新增點位

        #region 讀取資料庫寫入Shp

        /// <summary>
        /// 讀取資料庫寫入 Shp
        /// </summary>
        [Route("get/exportShp")]
        public void GetExportShp()
        {
            // 初始化
            string direPath = @"D:\shapefile";
            GdalConfiguration.ConfigureGdal();
            GdalConfiguration.ConfigureOgr();
            Gdal.SetConfigOption("GDAL_FILENAME_IS_UTF8", "YES");
            Gdal.SetConfigOption("SHAPE_ENCODING", "");
            Gdal.AllRegister();
            Ogr.RegisterAll();
            pDriver = Ogr.GetDriverByName("ESRI Shapefile");
            SpatialReference sr = new SpatialReference("");

            // 創建圖層
            Layer pLayer = CreateLayer(direPath, wkbGeometryType.wkbPoint, sr);

            // 創建欄位
            FieldDefn pFieldDefn = new FieldDefn("Id", FieldType.OFTInteger);
            pFieldDefn.SetWidth(200);
            pLayer.CreateField(pFieldDefn, 1);

            pFieldDefn = new FieldDefn("Name", FieldType.OFTString);
            pFieldDefn.SetWidth(200);
            pLayer.CreateField(pFieldDefn, 1);

            pFieldDefn = new FieldDefn("Food", FieldType.OFTString);
            pFieldDefn.SetWidth(200);
            pLayer.CreateField(pFieldDefn, 1);

            pFieldDefn = new FieldDefn("Address", FieldType.OFTString);
            pFieldDefn.SetWidth(200);
            pLayer.CreateField(pFieldDefn, 1);

            pFieldDefn = new FieldDefn("Phone", FieldType.OFTString);
            pFieldDefn.SetWidth(200);
            pLayer.CreateField(pFieldDefn, 1);

            pFieldDefn = new FieldDefn("Lat", FieldType.OFTReal);
            pFieldDefn.SetWidth(200);
            pLayer.CreateField(pFieldDefn, 1);

            pFieldDefn = new FieldDefn("Longitude", FieldType.OFTReal);
            pFieldDefn.SetWidth(200);
            pLayer.CreateField(pFieldDefn, 1);

            // 獲得欄位索引值來新增資料
            FeatureDefn pFeatureDefn = pLayer.GetLayerDefn();
            int fieldId = pFeatureDefn.GetFieldIndex("Id");
            int fieldName = pFeatureDefn.GetFieldIndex("Name");
            int fieldFood = pFeatureDefn.GetFieldIndex("Food");
            int fieldAddress = pFeatureDefn.GetFieldIndex("Address"); ;
            int fieldPhone = pFeatureDefn.GetFieldIndex("Phone");
            int fieldLat = pFeatureDefn.GetFieldIndex("Lat");
            int fieldLongitude = pFeatureDefn.GetFieldIndex("Longitude");

            // 獲取資料庫資料後 / 插入要素
            var conStr = @" SELECT * FROM dbo.Info";
            var dataset = sqlConn.DbQuery(conStr);

            Feature pFeature = new Feature(pFeatureDefn);
            foreach (var data in dataset)
            {
                Geometry geometry = Geometry.CreateFromWkt($"POINT({data.Lat} {data.Longitude})");
                pFeature.SetGeometry(geometry);
                pFeature.SetField(fieldId, data.Id);
                pFeature.SetField(fieldName, data.Name);
                pFeature.SetField(fieldFood, data.Food);
                pFeature.SetField(fieldAddress, data.Address);
                pFeature.SetField(fieldPhone, data.Phone);
                pFeature.SetField(fieldLat, data.Lat);
                pFeature.SetField(fieldLongitude, data.Longitude);
                pLayer.CreateFeature(pFeature);
                pLayer.SetFeature(pFeature);
            }
            pDriver.Register();
            dataSource.FlushCache();
            dataSource.Dispose();
            pLayer.Dispose();
            pDriver.Dispose();
        }

        #endregion 讀取資料庫寫入Shp

        #region 上傳及下載檔案(.kml/.shp/.ods)

        /// <summary>
        /// 讀取資料庫寫入 Kml
        /// </summary>
        [Route("get/downloadKml")]
        [HttpGet]
        public HttpResponseMessage DownloadKml()
        {
            string outputPath = $@"L:\New.kml";
            XmlTextWriter kml = new XmlTextWriter(outputPath, Encoding.UTF8);
            kml.WriteStartDocument();
            kml.WriteStartElement("kml", "http://www.opengis.net/kml/2.2"); //kml
            kml.WriteStartElement("Document"); //Document
            kml.WriteAttributeString("id", "root_doc");
            kml.WriteStartElement("Schema");
            kml.WriteAttributeString("name", "shapefile");
            kml.WriteAttributeString("id", "shapefile");
            kml.WriteStartElement("SimpleField");
            kml.WriteAttributeString("name", "Id");
            kml.WriteAttributeString("type", "float");
            kml.WriteEndElement(); // Id
            kml.WriteStartElement("SimpleField");
            kml.WriteAttributeString("name", "Food");
            kml.WriteAttributeString("type", "string");
            kml.WriteEndElement(); // Food
            kml.WriteStartElement("SimpleField");
            kml.WriteAttributeString("name", "Address");
            kml.WriteAttributeString("type", "string");
            kml.WriteEndElement(); // Address
            kml.WriteStartElement("SimpleField");
            kml.WriteAttributeString("name", "Phone");
            kml.WriteAttributeString("type", "string");
            kml.WriteEndElement(); // Phone
            kml.WriteStartElement("SimpleField");
            kml.WriteAttributeString("name", "Lat");
            kml.WriteAttributeString("type", "float");
            kml.WriteEndElement(); // Lat
            kml.WriteStartElement("SimpleField");
            kml.WriteAttributeString("name", "Longitude");
            kml.WriteAttributeString("type", "float");
            kml.WriteEndElement(); // Longitude
            kml.WriteEndElement(); // Schema
            kml.WriteStartElement("Folder"); // Folder
            kml.WriteStartElement("name");
            kml.WriteString("shapefile");
            kml.WriteEndElement(); // name

            // 寫入資料
            string conStr = @"SELECT * FROM dbo.Info";
            var dataset = sqlConn.DbQuery(conStr);
            foreach (var data in dataset)
            {
                kml.WriteStartElement("Placemark"); // Placemark
                kml.WriteStartElement("name"); // name
                kml.WriteEndElement(); // name
                kml.WriteStartElement("ExtendedData"); // ExtendedData
                kml.WriteStartElement("SchemaData"); // SchemaData
                kml.WriteAttributeString("schemaUrl", "#shapefile");
                foreach (var item in data)
                {
                    kml.WriteStartElement("SimpleData"); // SimpleData
                    kml.WriteAttributeString("name", item.Key);
                    kml.WriteString(item.Value.ToString());
                    kml.WriteEndElement();
                }
                kml.WriteEndElement(); // ExtendedData
                kml.WriteEndElement(); // SchemaData
                kml.WriteStartElement("Point");
                kml.WriteStartElement("coordinates");
                kml.WriteString($"{data.Lat},{data.Longitude}");
                kml.WriteEndElement(); // coordinates
                kml.WriteEndElement(); // Point
                kml.WriteEndElement(); // Placemark
            }
            kml.WriteEndElement(); // Folder
            kml.WriteEndElement(); // Documment
            kml.WriteEndElement(); // kml
            kml.WriteEndDocument();
            kml.Close();
            return FileResult(outputPath, "application/vnd.oasis.opendocument.spreadsheet");
        }

        /// <summary>
        /// 獲取 Kml 資料寫入資料庫
        /// </summary>
        [Route("post/uploadKml")]
        [HttpPost]
        public void UploadKml()
        {
            var request = HttpContext.Current.Request;
            var file = request.Files[0];
            string path = $@"L:\Temp\{file.FileName}";
            var document = XDocument.Load(path);
            var ns = document.Root.Name.Namespace;
            //get every placemark element in the document
            var placemarks = document.Descendants(ns + "Placemark");
            //loop through each placemark and separate it into coordinates and bearings
            string conStr = "INSERT INTO info (Name, Food, Address, Phone, Lat, Longitude, Location) VALUES(@Name, @Food, @Address, @Phone, @Lat, @Longitude, @Location)";
            var data = new List<Info>();
            foreach (var point in placemarks)
            {
                string coordinate = point.Descendants(ns + "coordinates").First().Value;
                string[] coordinateArray = coordinate.Split(",");
                List<XElement> pointData = point.Descendants(ns + "SimpleData").ToList();
                data.Add(new Info()
                {
                    Name = point.Descendants(ns + "name").First().Value,
                    Food = (string)pointData[1],
                    Address = (string)pointData[2],
                    Phone = (string)pointData[3],
                    Lat = SqlGeometry.STGeomFromText(new SqlChars($"POINT ({coordinateArray[0]} {coordinateArray[1]} )"), 4326).STX.Value,
                    Longitude = SqlGeometry.STGeomFromText(new SqlChars($"POINT ({coordinateArray[0]} {coordinateArray[1]} )"), 4326).STY.Value,
                    Location = SqlGeometry.STGeomFromText(new SqlChars($"POINT ({coordinateArray[0]} {coordinateArray[1]} )"), 4326)
                });
            }
            sqlConn.DbExecute(conStr, data);
        }

<<<<<<< HEAD
=======
        ///// <summary>
        ///// 讀取資料庫寫入 Kml
        ///// </summary>
        //[Route("get/downloadKml")]
        //[HttpGet]
        //public HttpResponseMessage DownloadKml()
        //{
        //}

>>>>>>> 327c7fe82d91496ca66b6d98d02b7d8f90675db2
        /// <summary>
        /// 獲取 Shp 資料寫入資料庫
        /// </summary>
        [Route("post/uploadShp")]
        [HttpPost]
        public void UploadShp()
        {
            var request = HttpContext.Current.Request;
            var file = request.Files[0];
            string path = $@"L:\Temp\{file.FileName}";
            m_Shp.InitinalGdal();
            // 数据源
            Driver pDriver = Ogr.GetDriverByName("ESRI Shapefile");
            DataSource pDataSource = pDriver.Open(path, 0);
            Layer pLayer = pDataSource.GetLayerByName(Path.GetFileNameWithoutExtension(path));
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
            string conStr = "INSERT INTO info (Name, Food, Address, Phone, Lat, Longitude, Location) VALUES(@Name, @Food, @Address, @Phone, @Lat, @Longitude, @Location)";
            var data = new List<Info>();
            int listDicLength = listDic.Count;
            for (int i = 0; i < listDicLength; i++)
            {
                data.Add(new Info()
                {
                    Name = listDic[i]["Name"].ToString(),
                    Food = listDic[i]["Food"].ToString(),
                    Address = listDic[i]["Address"].ToString(),
                    Phone = listDic[i]["Phone"].ToString(),
                    Lat = SqlGeometry.STGeomFromText(new SqlChars(listDic[i]["Geo"].ToString()), 4326).STX.Value,
                    Longitude = SqlGeometry.STGeomFromText(new SqlChars(listDic[i]["Geo"].ToString()), 4326).STY.Value,
                    Location = SqlGeometry.STGeomFromText(new SqlChars(listDic[i]["Geo"].ToString()), 4326)
                });
            }
            sqlConn.DbExecute(conStr, data);
            pLayer.Dispose();
            pDriver.Dispose();
        }

        /// <summary>
        /// 讀取資料庫寫入 Shp
        /// </summary>
        [Route("get/downloadShp")]
        [HttpGet]
        public HttpResponseMessage DownloadShp()
        {
            // 初始化
            string direPath = $@"L:\shapefile";
            string fileName = $@"L:\shapefile\shapefile.shp";
            if (Directory.Exists(direPath))
            {
                Directory.Delete(direPath, true);
            }
            GdalConfiguration.ConfigureGdal();
            GdalConfiguration.ConfigureOgr();
            Gdal.SetConfigOption("GDAL_FILENAME_IS_UTF8", "YES");
            Gdal.SetConfigOption("SHAPE_ENCODING", "");
            Gdal.AllRegister();
            Ogr.RegisterAll();
            pDriver = Ogr.GetDriverByName("ESRI Shapefile");
            SpatialReference sr = new SpatialReference("");

            // 創建圖層
            Layer pLayer = CreateLayer(direPath, wkbGeometryType.wkbPoint, sr);

            // 創建欄位
            FieldDefn pFieldDefn = new FieldDefn("Id", FieldType.OFTInteger);
            pFieldDefn.SetWidth(200);
            pLayer.CreateField(pFieldDefn, 1);

            pFieldDefn = new FieldDefn("Name", FieldType.OFTString);
            pFieldDefn.SetWidth(200);
            pLayer.CreateField(pFieldDefn, 1);

            pFieldDefn = new FieldDefn("Food", FieldType.OFTString);
            pFieldDefn.SetWidth(200);
            pLayer.CreateField(pFieldDefn, 1);

            pFieldDefn = new FieldDefn("Address", FieldType.OFTString);
            pFieldDefn.SetWidth(200);
            pLayer.CreateField(pFieldDefn, 1);

            pFieldDefn = new FieldDefn("Phone", FieldType.OFTString);
            pFieldDefn.SetWidth(200);
            pLayer.CreateField(pFieldDefn, 1);

            pFieldDefn = new FieldDefn("Lat", FieldType.OFTReal);
            pFieldDefn.SetWidth(200);
            pLayer.CreateField(pFieldDefn, 1);

            pFieldDefn = new FieldDefn("Longitude", FieldType.OFTReal);
            pFieldDefn.SetWidth(200);
            pLayer.CreateField(pFieldDefn, 1);

            // 獲得欄位索引值來新增資料
            FeatureDefn pFeatureDefn = pLayer.GetLayerDefn();
            int fieldId = pFeatureDefn.GetFieldIndex("Id");
            int fieldName = pFeatureDefn.GetFieldIndex("Name");
            int fieldFood = pFeatureDefn.GetFieldIndex("Food");
            int fieldAddress = pFeatureDefn.GetFieldIndex("Address"); ;
            int fieldPhone = pFeatureDefn.GetFieldIndex("Phone");
            int fieldLat = pFeatureDefn.GetFieldIndex("Lat");
            int fieldLongitude = pFeatureDefn.GetFieldIndex("Longitude");

            // 獲取資料庫資料後 / 插入要素
            var conStr = @" SELECT * FROM dbo.Info";
            var dataset = sqlConn.DbQuery(conStr);

            Feature pFeature = new Feature(pFeatureDefn);
            foreach (var data in dataset)
            {
                Geometry geometry = Geometry.CreateFromWkt($"POINT({data.Lat} {data.Longitude})");
                pFeature.SetGeometry(geometry);
                pFeature.SetField(fieldId, data.Id);
                pFeature.SetField(fieldName, data.Name);
                pFeature.SetField(fieldFood, data.Food);
                pFeature.SetField(fieldAddress, data.Address);
                pFeature.SetField(fieldPhone, data.Phone);
                pFeature.SetField(fieldLat, data.Lat);
                pFeature.SetField(fieldLongitude, data.Longitude);
                pLayer.CreateFeature(pFeature);
                pLayer.SetFeature(pFeature);
            }
            pDriver.Register();
            dataSource.FlushCache();
            dataSource.Dispose();
            pLayer.Dispose();
            pDriver.Dispose();
            return FileResult(fileName, "application/vnd.oasis.opendocument.spreadsheet");
        }

        /// <summary>
        /// 上傳 ods
        /// </summary>
        [Route("post/uploadOds")]
        [HttpPost]
        public void UploadOds()
        {
            var request = HttpContext.Current.Request;
            string path;
            if (request.Files.Count > 0)
            {
                var file = request.Files[0];
                path = $@"L:\Temp\{file.FileName}";
                var row = 1;
                using (var calc = new Calc(path))
                {
                    var workSheet = calc.Tables[1];
                    int columnsLength = workSheet.ColumnCount - 2;
                    int rowsLength = workSheet.RowCount;
                    string conStr = @"INSERT INTO Info (Name, Food, Address, Phone, Lat, Longitude, Location) VALUES(@Name, @Food, @Address, @Phone, @Lat, @Longitude, @Location)";
                    var data = new List<Info>();
                    for (int i = 0; i < rowsLength - 1; i++)
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
                }
            }
        }

        /// <summary>
        /// 下載 Ods 檔案
        /// </summary>
        /// <returns></returns>
        [Route("get/downloadOds")]
        [HttpGet]
        public HttpResponseMessage DownloadOds()
        {
            var path = @"L:\New.ods";
            if (Directory.Exists(path))
            {
                Directory.Delete(path, true);
            }
            using (var calc = new Calc())
            {
                try
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
                            column++;
                        }
                        row++;
                        column = 0;
                    }
                    calc.SaveAs(path);
                    calc.Close();
                }
                catch (Exception e)
                {
                }
            }
            return FileResult(path, "application/vnd.oasis.opendocument.spreadsheet");
        }

        #endregion 上傳及下載檔案(.kml/.shp/.ods)

        #region 修改檔案

        /// <summary>
        /// 修改本地 Ods 檔案
        /// </summary>
        [Route("get/ods")]
        public void Get()
        {
            Logger _logger = LogManager.GetCurrentClassLogger();
            Calc.LogPath = @"D:\log";
            var path = @"D:\test.ods";
            var outputPath = @"D:\output.ods";
            using (var calc = new Calc(path))
            {
                try
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
                catch (Exception e)
                {
                    _logger.Error("[Error in ClientController.Edit - id: " + " - Error: " + e.Message + "]");
                    /// add redirect link here
                }
            }
        }

        #endregion 修改檔案
    }
}