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
using DataSource = OSGeo.OGR.DataSource;
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
using HttpRequest = System.Web.HttpRequest;
using System.Web.Routing;
using unoidl.com.sun.star.awt;
using System.Web.WebPages;
using System.IO.Compression;
using System.Net.Http.Formatting;
using System.Web.Hosting;
using System.Diagnostics;
using System.Data.Services.Client;
using System.Runtime.InteropServices.ComTypes;
using System.Web.UI.WebControls;
using Table = OpenDocumentLib.sheet.Table;
using System.Windows.Shapes;
using Path = System.IO.Path;
using Line = OpenDocumentLib.sheet.Line;
using File = System.IO.File;

namespace FileTest.Controllers
{
    [EnableCors(origins: "*", headers: "*", methods: "*", SupportsCredentials = true)]
    public class OdsController : ApiController
    {
        public ShpRead m_Shp;
        public static Driver pDriver;
        public static DataSource dataSource;

        // 根目錄
        private static readonly string _rootFolderName = HttpContext.Current.Server.MapPath($"~");

        // Shp的資料夾
        private static readonly string _shpFolderName = HttpContext.Current.Server.MapPath($"~/Files/Shp/");

        // 處理的暫存資料夾
        private static readonly string _tempFolderName = HttpContext.Current.Server.MapPath($"~/Files/Temp/");

        //shp
        [DllImport("gdal204.dll", EntryPoint = "OGR_F_GetFieldAsString", CallingConvention = CallingConvention.Cdecl)]
        public static extern System.IntPtr OGR_F_GetFieldAsString(HandleRef handle, int index);

        public OdsController()
        {
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
            pFeature.SetField(fieldIndex_B, "價格B");
            pFeature.SetField(fieldIndex_C, "北平路");
            pLayer.CreateFeature(pFeature);

            geometry = Geometry.CreateFromWkt("POINT(11 110)");
            pFeature.SetGeometry(geometry);
            pFeature.SetField(fieldIndex_A, "3");
            pFeature.SetField(fieldIndex_B, "價格C");
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
            string direPath = @"C:\shapefile";
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

            pFieldDefn = new FieldDefn("Category", FieldType.OFTString);
            pFieldDefn.SetWidth(200);
            pLayer.CreateField(pFieldDefn, 1);

            pFieldDefn = new FieldDefn("Price", FieldType.OFTString);
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
            int fieldCategory = pFeatureDefn.GetFieldIndex("Category");
            int fieldPrice = pFeatureDefn.GetFieldIndex("Price");
            int fieldAddress = pFeatureDefn.GetFieldIndex("Address"); ;
            int fieldPhone = pFeatureDefn.GetFieldIndex("Phone");
            int fieldLat = pFeatureDefn.GetFieldIndex("Lat");
            int fieldLongitude = pFeatureDefn.GetFieldIndex("Longitude");

            // 獲取資料庫資料後 / 插入要素
            var conStr = @" SELECT * FROM dbo.Info";
            var dataset = SqlConn.DbQuery(conStr);

            Feature pFeature = new Feature(pFeatureDefn);
            foreach (var data in dataset)
            {
                Geometry geometry = Geometry.CreateFromWkt($"POINT({data.Lat} {data.Longitude})");
                pFeature.SetGeometry(geometry);
                pFeature.SetField(fieldId, data.Id);
                pFeature.SetField(fieldCategory, data.Category);
                pFeature.SetField(fieldPrice, data.Price);
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
        [Route("get/downloadKml/{ids}")]
        [HttpGet]
        public HttpResponseMessage DownloadKml(string ids)
        {
            string outputPath = $@"C:\files\download\download.kml";
            XmlTextWriter kml = new XmlTextWriter(outputPath, Encoding.UTF8);
            kml.WriteStartDocument();
            kml.WriteStartElement("kml", "http://www.opengis.net/kml/2.2"); //kml
            kml.WriteStartElement("Document"); //Document
            kml.WriteAttributeString("id", "root_doc");
            kml.WriteStartElement("Schema");
            kml.WriteAttributeString("name", "shapefile");
            kml.WriteAttributeString("id", "shapefile");
            kml.WriteStartElement("SimpleField");
            kml.WriteAttributeString("name", "Category");
            kml.WriteAttributeString("type", "string");
            kml.WriteEndElement(); // Category
            kml.WriteStartElement("SimpleField");
            kml.WriteAttributeString("name", "Price");
            kml.WriteAttributeString("type", "string");
            kml.WriteEndElement(); // Price
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
            string conStr = $@"SELECT * FROM dbo.Info WHERE id in ({ids})";
            var dataset = SqlConn.DbQuery(conStr);
            foreach (var data in dataset)
            {
                kml.WriteStartElement("Placemark"); // Placemark
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
            kml.Dispose();
            return FileResult(outputPath, "application/vnd.oasis.opendocument.spreadsheet");
        }

        /// <summary>
        /// 獲取 Kml 資料寫入資料庫
        /// </summary>
        [Route("post/uploadKml")]
        [HttpPost]
        public void UploadKml()
        {
            HttpRequest request = HttpContext.Current.Request;
            string readPath = $@"C:\files\getUpload";
            if (request.Files.Count > 0)
            {
                IEnumerable<HttpPostedFile> files = Enumerable.Range(0, request.Files.Count).
            Select(i => request.Files[i]);
                foreach (var file in files)
                {
                    // 上傳位置
                    if (!Directory.Exists(readPath))
                        Directory.CreateDirectory(readPath);
                    if (System.IO.File.Exists(readPath)) System.IO.File.Delete(readPath);
                    string path = Path.Combine(readPath, file.FileName);
                    file.SaveAs(path);
                    var document = XDocument.Load(path);
                    var ns = document.Root.Name.Namespace;
                    var placemarks = document.Descendants(ns + "Placemark");
                    string conStr = "INSERT INTO info (Category, Price, Address, Phone, Lat, Longitude, Location) VALUES(@Category, @Price, @Address, @Phone, @Lat, @Longitude, @Location)";
                    var data = new List<Info>();
                    foreach (var point in placemarks)
                    {
                        string coordinate = point.Descendants(ns + "coordinates").First().Value;
                        string[] coordinateArray = coordinate.Split(",");
                        List<XElement> pointData = point.Descendants(ns + "SimpleData").ToList();
                        data.Add(new Info()
                        {
                            Category = (string)pointData[1],
                            Price = (int)pointData[2],
                            Address = (string)pointData[3],
                            Phone = (string)pointData[4],
                            Lat = (double)pointData[5],
                            Longitude = (double)pointData[6],
                            Location = SqlGeometry.STGeomFromText(new SqlChars($"POINT ({coordinateArray[0]} {coordinateArray[1]} )"), 4326)
                        });
                    }
                    SqlConn.DbExecute(conStr, data);
                }
            }
        }

        /// <summary>
        /// 獲取 Shp 資料寫入資料庫
        /// </summary>
        [Route("post/uploadShp")]
        [HttpPost]
        public void UploadShp()
        {
            HttpRequest request = HttpContext.Current.Request;
            string readPath = $@"C:\files\getUpload";
            if (request.Files.Count > 0)
            {
                string shapefile = string.Empty;
                IEnumerable<HttpPostedFile> files = Enumerable.Range(0, request.Files.Count).
            Select(i => request.Files[i]);
                foreach (var file in files)
                {
                    // 上傳位置
                    if (!Directory.Exists(readPath))
                        Directory.CreateDirectory(readPath);
                    if (System.IO.File.Exists(readPath)) System.IO.File.Delete(readPath);
                    string path = Path.Combine(readPath, file.FileName);
                    file.SaveAs(path);
                    if (Path.GetExtension(file.FileName).ToLower() == ".shp")
                        shapefile = Path.Combine(readPath, file.FileName);
                }
                m_Shp.InitinalGdal();
                // 數據源
                Driver pDriver = Ogr.GetDriverByName("ESRI Shapefile");
                DataSource pDataSource = pDriver.Open(shapefile, 0);
                Layer pLayer = pDataSource.GetLayerByName(Path.GetFileNameWithoutExtension(shapefile));
                pLayer.GetSpatialRef();
                // 字段結構
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
                pDriver.Dispose();
                string conStr = "INSERT INTO info (Category, Price, Address, Phone, Lat, Longitude, Location) VALUES(@Category, @Price, @Address, @Phone, @Lat, @Longitude, @Location)";
                var data = new List<Info>();
                int listDicLength = listDic.Count;
                for (int i = 0; i < listDicLength; i++)
                {
                    data.Add(new Info()
                    {
                        Category = listDic[i]["Category"].ToString(),
                        Price = Convert.ToInt32(listDic[i]["Price"]),
                        Address = listDic[i]["Address"].ToString(),
                        Phone = listDic[i]["Phone"].ToString(),
                        Lat = SqlGeometry.STGeomFromText(new SqlChars(listDic[i]["Geo"].ToString()), 4326).STX.Value,
                        Longitude = SqlGeometry.STGeomFromText(new SqlChars(listDic[i]["Geo"].ToString()), 4326).STY.Value,
                        Location = SqlGeometry.STGeomFromText(new SqlChars(listDic[i]["Geo"].ToString()), 4326)
                    });
                }
                SqlConn.DbExecute(conStr, data);
                pLayer.Dispose();
                pDriver.Dispose();
            }
        }

        /// <summary>
        /// 讀取資料庫寫入 Shp
        /// </summary>
        [Route("get/downloadShp/{ids}")]
        [HttpGet]
        public HttpResponseMessage DownloadShp(string ids)
        {
            // 初始化
            string direPath = $@"C:\files\download\shapefile";
            string zipPath = $@"C:\files\download\shapefile.zip";
            string outputPath = $@"C:\files\download\shapefile\shapefile.shp";
            if (Directory.Exists(direPath))
            {
                Directory.Delete(direPath, true);
                Directory.CreateDirectory(direPath);
            }
            else
            {
                Directory.CreateDirectory(direPath);
            }
            if (File.Exists(zipPath))
            {
                File.Delete(zipPath);
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

            pFieldDefn = new FieldDefn("Category", FieldType.OFTString);
            pFieldDefn.SetWidth(200);
            pLayer.CreateField(pFieldDefn, 1);

            pFieldDefn = new FieldDefn("Price", FieldType.OFTString);
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
            int fieldCategory = pFeatureDefn.GetFieldIndex("Category");
            int fieldPrice = pFeatureDefn.GetFieldIndex("Price");
            int fieldAddress = pFeatureDefn.GetFieldIndex("Address"); ;
            int fieldPhone = pFeatureDefn.GetFieldIndex("Phone");
            int fieldLat = pFeatureDefn.GetFieldIndex("Lat");
            int fieldLongitude = pFeatureDefn.GetFieldIndex("Longitude");

            // 獲取資料庫資料後 / 插入要素
            string conStr = $@"SELECT * FROM dbo.Info WHERE id in ({ids})";
            var dataset = SqlConn.DbQuery(conStr);

            Feature pFeature = new Feature(pFeatureDefn);
            foreach (var data in dataset)
            {
                Geometry geometry = Geometry.CreateFromWkt($"POINT({data.Lat} {data.Longitude})");
                pFeature.SetGeometry(geometry);
                pFeature.SetField(fieldId, data.Id);
                pFeature.SetField(fieldCategory, data.Category);
                pFeature.SetField(fieldPrice, data.Price);
                pFeature.SetField(fieldAddress, data.Address);
                pFeature.SetField(fieldPhone, data.Phone);
                pFeature.SetField(fieldLat, data.Lat);
                pFeature.SetField(fieldLongitude, data.Longitude);
                pLayer.CreateFeature(pFeature);
                pLayer.SetFeature(pFeature);
            }
            pFeature.Dispose();
            pFieldDefn.Dispose();
            pLayer.Dispose();
            pDriver.Dispose();
            sr.Dispose();
            dataSource.FlushCache();
            dataSource.Dispose();
            ZipFile.CreateFromDirectory(direPath, zipPath);
            return FileResult(zipPath, "application/vnd.oasis.opendocument.spreadsheet");
        }

        /// <summary>
        /// 上傳 ods
        /// </summary>
        [Route("post/uploadOds")]
        [HttpPost]
        public void UploadOds()
        {
            HttpRequest request = HttpContext.Current.Request;
            string readPath = $@"C:\files\getUpload";
            if (request.Files.Count > 0)
            {
                IEnumerable<HttpPostedFile> files = Enumerable.Range(0, request.Files.Count)
            .Select(i => request.Files[i]);
                foreach (var file in files)
                {
                    // 上傳位置
                    if (!Directory.Exists(readPath))
                        Directory.CreateDirectory(readPath);
                    if (System.IO.File.Exists(readPath)) System.IO.File.Delete(readPath);
                    string path = Path.Combine(readPath, file.FileName);
                    file.SaveAs(path);
                    int row = 1;
                    using (Calc calc = new Calc(path))
                    {
                        try
                        {
                            Table workSheet = calc.Tables[0];
                            int columnsLength = workSheet.ColumnCount - 2;
                            int rowsLength = workSheet.RowCount;
                            string conStr = @"INSERT INTO Info (Category, Price, Address, Phone, Lat, Longitude, Location) VALUES(@Category, @Price, @Address, @Phone, @Lat, @Longitude, @Location)";
                            List<Info> infoData = new List<Info>();
                            for (int i = 0; i < rowsLength - 1; i++)
                            {
                                double Lat = workSheet[row, 5].Formula.ToDouble();
                                double Longitude = workSheet[row, 6].Formula.ToDouble();
                                infoData.Add(new Info()
                                {
                                    Category = workSheet[row, 1].Formula,
                                    Price = workSheet[row, 2].Formula.ToInt32(),
                                    Address = workSheet[row, 3].Formula,
                                    Phone = workSheet[row, 4].Formula,
                                    Lat = Lat,
                                    Longitude = Longitude,
                                    Location = SqlGeometry.Point(Lat, Longitude, 4326)
                                });
                                row++;
                            }
                            SqlConn.DbExecute(conStr, infoData);
                        }
                        catch
                        {
                            throw new Exception("can not upload .");
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 下載 Ods 檔案
        /// </summary>
        /// <returns></returns>
        [Route("get/downloadOds/{ids}")]
        [HttpGet]
        public HttpResponseMessage DownloadOds(string ids)
        {
            string outputPath = @"C:\files\download\download.ods";

            using (var calc = new Calc())
            {
                try
                {
                    Table workSheet = calc.Tables[0];
                    Font headerFont = new Font("標楷體", 28, FontStyle.Bold),
                    colFont = new Font("標楷體", 14, FontStyle.Bold),
                    rowFont = new Font("標楷體", 12);
                    Line line = new Line() { Color = Color.Black, OuterWidth = 20 };
                    int header = 0;
                    int row = 0;
                    int column = 0;
                    string conStr = $@"SELECT * FROM dbo.Info WHERE id in ({ids})";
                    List<dynamic> dataset = SqlConn.DbQuery(conStr);
                    foreach (var rowData in dataset)
                    {
                        foreach (var field in rowData)
                        {
                            if (row <= header)
                            {
                                workSheet[row, column].Formula = field.Key;
                            }
                            workSheet[row + 1, column].Formula = Convert.ToString(field.Value);
                            column++;
                        }
                        row++;
                        column = 0;
                    }
                    calc.SaveAs(outputPath);
                }
                catch
                {
                    throw new Exception("can not download .");
                }
            }
            return FileResult(outputPath, "application/vnd.oasis.opendocument.spreadsheet");
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
            Calc.LogPath = @"C:\log";
            var path = @"C:\test.ods";
            var outputPath = @"C:\output.ods";
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

        #region 我的練習

        [Route("get/myTest")]
        [HttpGet]
        public void Test()
        {
            List<string> test = new List<string>() { "aa", "bb", "cc", "aAA", "cCC", "d" };
            List<string> copyTest = new List<string>(test);
            Dictionary<string, string> newTest = new Dictionary<string, string>();
            int copyTestLength = copyTest.Count;
            var sum = new StringBuilder();
            while (copyTest.Count != 0)
            {
                List<string> equalList = copyTest.Where(p => p.Contains(copyTest[0][0])).ToList();
                foreach (var item in equalList)
                {
                    sum.Append(item);
                    copyTest.Remove(item);
                }
                newTest.Add(sum[0].ToString(), sum.ToString());
                sum.Clear();
            }
        }

        #endregion 我的練習
    }
}