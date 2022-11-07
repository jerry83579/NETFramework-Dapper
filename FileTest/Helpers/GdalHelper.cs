using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OSGeo.GDAL;
using OSGeo.OGR;
using OSGeo.OSR;
using System.IO;
using FileTest;

namespace ShpUploadAPI
{
    public class GdalHelper
    {
        [DllImport("gdal204.dll", EntryPoint = "OGR_F_GetFieldAsString", CallingConvention = CallingConvention.Cdecl)]
        public static extern System.IntPtr OGR_F_GetFieldAsString(HandleRef handle, int index);

        /// <summary>
        /// 获取要素图层
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static OSGeo.OGR.Layer GetLayer(string filePath)
        {
            // 注册GDAL
            OSGeo.GDAL.Gdal.SetConfigOption("GDAL_FILENAME_IS_UTF8", "YES");
            OSGeo.GDAL.Gdal.SetConfigOption("SHAPE_ENCODING", "");
            OSGeo.GDAL.Gdal.AllRegister();
            OSGeo.OGR.Ogr.RegisterAll();

            // 获取图层
            OSGeo.OGR.Driver pDriver = Ogr.GetDriverByName("ESRI Shapefile");
            OSGeo.OGR.DataSource pDataSource = pDriver.Open(filePath, 1);
            OSGeo.OGR.Layer pLayer = pDataSource.GetLayerByName(System.IO.Path.GetFileNameWithoutExtension(filePath));
            return pLayer;
        }

        /// <summary>
        /// 创建要素图层
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="geometryType"></param>
        /// <param name="spatialReference"></param>
        /// <returns></returns>
        public static OSGeo.OGR.Layer CreateLayer(string filePath, OSGeo.OGR.wkbGeometryType geometryType, OSGeo.OSR.SpatialReference spatialReference)
        {
            // 注册GDAL
            GdalConfiguration.ConfigureGdal();
            GdalConfiguration.ConfigureOgr();
            OSGeo.GDAL.Gdal.SetConfigOption("GDAL_FILENAME_IS_UTF8", "YES");
            OSGeo.GDAL.Gdal.SetConfigOption("SHAPE_ENCODING", "");
            OSGeo.GDAL.Gdal.AllRegister();
            OSGeo.OGR.Ogr.RegisterAll();

            // 数据源
            OSGeo.OGR.Driver pDriver = Ogr.GetDriverByName("ESRI Shapefile");
            OSGeo.OGR.DataSource pDataSource = pDriver.CreateDataSource(filePath, null);
            OSGeo.OGR.Layer pLayer = null;

            // 创建图层
            if (geometryType == wkbGeometryType.wkbPoint)
            {
                pLayer = pDataSource.CreateLayer(System.IO.Path.GetFileNameWithoutExtension(filePath), spatialReference, wkbGeometryType.wkbPoint, null);
            }
            else if (geometryType == wkbGeometryType.wkbLineString)
            {
                pLayer = pDataSource.CreateLayer(System.IO.Path.GetFileNameWithoutExtension(filePath), spatialReference, wkbGeometryType.wkbLineString, null);
            }
            else
            {
                pLayer = pDataSource.CreateLayer(System.IO.Path.GetFileNameWithoutExtension(filePath), spatialReference, wkbGeometryType.wkbPolygon, null);
            }
            return pLayer;
        }

        /// <summary>
        /// 读取要素图层属性表
        /// </summary>
        /// <param name="filePath"></param>
        public static List<Dictionary<string, object>> ReadLayerAttributes(string filePath)
        {
            // 注册GDAL
            OSGeo.GDAL.Gdal.SetConfigOption("GDAL_FILENAME_IS_UTF8", "YES");
            OSGeo.GDAL.Gdal.SetConfigOption("SHAPE_ENCODING", "");
            OSGeo.GDAL.Gdal.AllRegister();
            OSGeo.OGR.Ogr.RegisterAll();

            // 数据源
            OSGeo.OGR.Driver pDriver = OSGeo.OGR.Ogr.GetDriverByName("ESRI Shapefile");
            OSGeo.OGR.DataSource pDataSource = pDriver.Open(filePath, 0);
            OSGeo.OGR.Layer pLayer = pDataSource.GetLayerByName(System.IO.Path.GetFileNameWithoutExtension(filePath));
            pLayer.GetSpatialRef();
            // 字段结构
            OSGeo.OGR.Feature pFeature = pLayer.GetNextFeature();
            OSGeo.OGR.FeatureDefn pFeatureDefn = pLayer.GetLayerDefn();
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
                    OSGeo.OGR.FieldDefn pFieldDefn = pFeature.GetFieldDefnRef(i);
                    string fieldName = pFieldDefn.GetName();
                    switch (pFieldDefn.GetFieldType())
                    {
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
                                dicData.Add(fieldName, Marshal.PtrToStringAnsi(pIntPtr));
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
            pDriver.DeleteDataSource(filePath);
            pDriver.Dispose();
            return listDic;
        }

        /// <summary>
        /// 插入要素
        /// </summary>
        /// <param name="filePath"></param>
        public static void InsertFeatures(string filePath)
        {
            // 注册GDAL
            OSGeo.GDAL.Gdal.SetConfigOption("GDAL_FILENAME_IS_UTF8", "YES");
            OSGeo.GDAL.Gdal.SetConfigOption("SHAPE_ENCODING", "");
            OSGeo.GDAL.Gdal.AllRegister();
            OSGeo.OGR.Ogr.RegisterAll();

            // 数据源
            OSGeo.OGR.Driver pDriver = OSGeo.OGR.Ogr.GetDriverByName("ESRI Shapefile");
            OSGeo.OGR.DataSource pDataSource = pDriver.Open(filePath, 1);
            OSGeo.OGR.Layer pLayer = pDataSource.GetLayerByName(System.IO.Path.GetFileNameWithoutExtension(filePath));

            // 字段索引
            OSGeo.OGR.FeatureDefn pFeatureDefn = pLayer.GetLayerDefn();
            int fieldIndex_A = pFeatureDefn.GetFieldIndex("A");
            int fieldIndex_B = pFeatureDefn.GetFieldIndex("B");
            int fieldIndex_C = pFeatureDefn.GetFieldIndex("C");

            // 插入要素
            OSGeo.OGR.Feature pFeature = new Feature(pFeatureDefn);
            OSGeo.OGR.Geometry geometry = Geometry.CreateFromWkt("POLYGON ((30 0,60 0,60 30,30 30,30 0))");
            pFeature.SetGeometry(geometry);
            pFeature.SetField(fieldIndex_A, 0);
            pFeature.SetField(fieldIndex_B, "211031");
            pFeature.SetField(fieldIndex_C, "张三");
            pLayer.CreateFeature(pFeature);

            geometry = Geometry.CreateFromWkt("POLYGON ((130 0,160 0,160 130,130 130,130 0))");
            pFeature.SetGeometry(geometry);
            pFeature.SetField(fieldIndex_A, 1);
            pFeature.SetField(fieldIndex_B, "211032");
            pFeature.SetField(fieldIndex_C, "李四");
            pLayer.CreateFeature(pFeature);
        }

        /// <summary>
        /// 創建欄位
        /// </summary>
        /// <param name="pLayer"></param>
        public static void CreateFields(OSGeo.OGR.Layer pLayer)
        {
            // int类型字段
            OSGeo.OGR.FieldDefn pFieldDefn = new OSGeo.OGR.FieldDefn("A", OSGeo.OGR.FieldType.OFTInteger);
            pLayer.CreateField(pFieldDefn, 1);

            // double类型字段
            pFieldDefn = new OSGeo.OGR.FieldDefn("B", OSGeo.OGR.FieldType.OFTReal);
            pFieldDefn.SetPrecision(3);
            pLayer.CreateField(pFieldDefn, 1);

            // string类型字段
            pFieldDefn = new OSGeo.OGR.FieldDefn("C", OSGeo.OGR.FieldType.OFTString);
            pFieldDefn.SetWidth(200);
            pLayer.CreateField(pFieldDefn, 1);
        }

        /// <summary>
        /// 输出字段名称
        /// </summary>
        /// <param name="pLayer"></param>
        public static void GetFieldsName(OSGeo.OGR.Layer pLayer)
        {
            OSGeo.OGR.FeatureDefn pFeatureDefn = pLayer.GetLayerDefn();
            for (int i = 0; i < pFeatureDefn.GetFieldCount(); i++)
            {
                OSGeo.OGR.FieldDefn pFieldDefn = pFeatureDefn.GetFieldDefn(i);
                Console.WriteLine(pFieldDefn.GetNameRef());
            }
        }

        /// <summary>
        /// 获取空间参考字符串
        /// </summary>
        /// <param name="pLayer"></param>
        /// <returns></returns>
        //public static string GetSpatialReferenceString(OSGeo.OGR.Layer pLayer)
        //{
        //    OSGeo.OSR.SpatialReference pSpatialReference = pLayer.GetSpatialRef();
        //    pSpatialReference.ExportToWkt(out string spatialReferenceString);
        //    return spatialReferenceString;
        //}

        /// <summary>
        /// 根据字符串生成空间参考
        /// </summary>
        /// <param name="spatialReferenceString"></param>
        /// <returns></returns>
        public static OSGeo.OSR.SpatialReference CreateSpatialReference(string spatialReferenceString)
        {
            OSGeo.OSR.SpatialReference pSpatialReference = new OSGeo.OSR.SpatialReference(spatialReferenceString);
            return pSpatialReference;
        }

        /// <summary>
        /// 座標轉換_測試中
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static string Reprojection(string filePath)
        {
            //.NET Core不支持gbk和gb2312，这里需要处理一下
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.InputEncoding = Encoding.UTF8;

            OSGeo.GDAL.Gdal.SetConfigOption("GDAL_FILENAME_IS_UTF8", "YES");
            OSGeo.GDAL.Gdal.SetConfigOption("SHAPE_ENCODING", "");
            OSGeo.GDAL.Gdal.AllRegister();
            OSGeo.OGR.Ogr.RegisterAll();

            OSGeo.OGR.Driver pDriver = OSGeo.OGR.Ogr.GetDriverByName("ESRI Shapefile");

            // input SpatialReference
            SpatialReference inSpatialRef = new OSGeo.OSR.SpatialReference("");
            inSpatialRef.ImportFromEPSG(3826);

            // output SpatialReference
            SpatialReference outSpatialRef = new OSGeo.OSR.SpatialReference("");
            outSpatialRef.ImportFromEPSG(4326);

            // create the CoordinateTransformation
            CoordinateTransformation coordTrans = new OSGeo.OSR.CoordinateTransformation(inSpatialRef, outSpatialRef);

            // get the input layer
            //DataSet inDataSet = driver.Open("c:\\data\\spatial\\basemap.shp",0);
            OSGeo.OGR.DataSource inDataSource = pDriver.Open(filePath, 0);
            OSGeo.OGR.Layer inLayer = inDataSource.GetLayerByIndex(0); //GetLayerByName(System.IO.Path.GetFileNameWithoutExtension(filePath));
                                                                       //Layer inLayer = inDataSet.GetLayer();

            // create the output layer
            //var outputShapefile = Path.Combine(Path.GetDirectoryName(filePath), Path.GetFileNameWithoutExtension(filePath) + "_new" + Path.GetExtension(filePath));
            //if (File.Exists(outputShapefile))
            //    pDriver.DeleteDataSource(outputShapefile);

            //OSGeo.OGR.Layer outLayer = CreateLayer(outputShapefile, inLayer.GetGeomType(), inLayer.GetSpatialRef());

            //CreateFields(outLayer);
            //InsertFeatures(outputShapefile);
            //OSGeo.OGR.DataSource outDataSource = pDriver.CreateDataSource(Path.GetDirectoryName(filePath),null);
            //outDataSource.CopyLayer(inLayer, Path.GetFileNameWithoutExtension(outputShapefile), null);

            //OSGeo.OGR.Layer outLayer = outDataSource.GetLayerByIndex(0);
            //Layer outLayer = outDataSource.GetLayerByName(Path.GetFileNameWithoutExtension(outputShapefile));

            //OSGeo.OSR.SpatialReference inSpatialReference = inLayer.GetSpatialRef();
            //inSpatialReference.ExportToWkt(out string inspatialReferenceString);

            //OSGeo.OSR.SpatialReference outSpatialReference = outLayer.GetSpatialRef();
            //outSpatialReference.ExportToWkt(out string outspatialReferenceString);

            //// add fields
            //OSGeo.OGR.FeatureDefn inLayerDefn = inLayer.GetLayerDefn();
            //for (int i = 0; i < inLayerDefn.GetFieldCount(); i++)
            //{
            //    FieldDefn fieldDefn = inLayerDefn.GetFieldDefn(i);
            //    outLayer.CreateField(fieldDefn, 1);
            //}

            //// get the output layer's feature definition
            //OSGeo.OGR.FeatureDefn outLayerDefn = outLayer.GetLayerDefn();

            // loop through the input features
            Feature inFeature = inLayer.GetNextFeature();
            while (inFeature != null)
            {
                // get the input geometry
                Geometry geom = inFeature.GetGeometryRef();

                geom.ExportToWkt(out string outspatialReferenceString);
                // reproject the geometry
                geom.Transform(coordTrans);

                geom.ExportToWkt(out string outspatialReferenceString2);
                // create a new feature
                //Feature outFeature = new Feature(outLayerDefn);
                //// set the geometry and attribute
                inFeature.SetGeometry(geom);
                //for (int i = 0; i < outLayerDefn.GetFieldCount(); i++)
                //    outFeature.SetField(outLayerDefn.GetFieldDefn(i).GetNameRef(), inFeature.GetFieldAsInteger(i));
                //// add the feature to the shapefile
                //outLayer.CreateFeature(outFeature);
                //// dereference the features and get the next input feature
                //outFeature = null;
                inFeature = inLayer.GetNextFeature();
            }
            // Save and close the shapefiles
            inDataSource = null;
            //outDataSource = null;
            return filePath;
        }
    }
}