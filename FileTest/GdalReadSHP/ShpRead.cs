using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Text;
//  
using System.IO;
using OSGeo.GDAL;
using OSGeo.OGR;
using OSGeo.OSR;
using System.Collections;
using FileTest;

namespace GdalReadSHP
{
    /// <summary>  
    /// 定義SHP解析類  
    /// </summary>  
    public class ShpRead
    {

        /// 保存SHP屬性字段  
        public OSGeo.OGR.Driver oDerive;
        public List<string> m_FeildList;
        private Layer oLayer;
        public string sCoordiantes;
        public ShpRead()
        {
            m_FeildList = new List<string>();
            oLayer = null;
            sCoordiantes = null;
        }

        /// <summary>  
        /// 初始化Gdal  
        /// </summary>  
        public OSGeo.OGR.Driver InitinalGdal()
        {
            GdalConfiguration.ConfigureGdal();
            GdalConfiguration.ConfigureOgr();
            Gdal.SetConfigOption("GDAL_FILENAME_IS_UTF8", "YES");
            Gdal.SetConfigOption("SHAPE_ENCODING", "");
            Gdal.AllRegister();
            Ogr.RegisterAll();

            oDerive = Ogr.GetDriverByName("ESRI Shapefile");
            if (oDerive == null)
            {
                //("文件不能打開，請檢查");
            }
            return oDerive;
        }

        /// <summary>  
        /// 獲取SHP文件的層  
        /// </summary>  
        /// <param name="sfilename"></param>  
        /// <param name="oLayer"></param>  
        /// <returns></returns>  
        public bool GetShpLayer(string sfilename)
        {
            if (null == sfilename || sfilename.Length <= 3)
            {
                oLayer = null;
                return false;
            }
            if (oDerive == null)
            {
                //("文件不能打開，請檢查");
            }
            DataSource ds = oDerive.Open(sfilename, 1);
            if (null == ds)
            {
                oLayer = null;
                return false;
            }
            int iPosition = sfilename.LastIndexOf("\\");
            string sTempName = sfilename.Substring(iPosition + 1, sfilename.Length - iPosition - 4 - 1);
            //.shp檔名為中文時
            //oLayer = ds.GetLayerByName(sTempName);
            oLayer = ds.GetLayerByIndex(0);
            if (oLayer == null)
            {
                ds.Dispose();
                return false;
            }
            return true;
        }
        /// <summary>  
        /// 獲取所有的屬性字段  
        /// </summary>  
        /// <returns></returns>  
        public bool GetFeilds()
        {
            if (null == oLayer)
            {
                return false;
            }
            m_FeildList.Clear();
            wkbGeometryType oTempGeometryType = oLayer.GetGeomType();
            List<string> TempstringList = new List<string>();

            //  
            FeatureDefn oDefn = oLayer.GetLayerDefn();
            int iFieldCount = oDefn.GetFieldCount();
            for (int iAttr = 0; iAttr < iFieldCount; iAttr++)
            {
                FieldDefn oField = oDefn.GetFieldDefn(iAttr);
                if (null != oField)
                {
                    m_FeildList.Add(oField.GetNameRef());
                }
            }
            return true;
        }
        /// <summary>  
        ///  獲取某條數據的字段內容  
        /// </summary>  
        /// <param name="iIndex"></param>  
        /// <param name="FeildStringList"></param>  
        /// <returns></returns>  
        public bool GetFeildContent(int iIndex, out List<string> FeildStringList)
        {
            FeildStringList = new List<string>();
            Feature oFeature = null;
            if ((oFeature = oLayer.GetFeature(iIndex)) != null)
            {

                FeatureDefn oDefn = oLayer.GetLayerDefn();
                int iFieldCount = oDefn.GetFieldCount();
                // 查找字段屬性  
                for (int iAttr = 0; iAttr < iFieldCount; iAttr++)
                {
                    FieldDefn oField = oDefn.GetFieldDefn(iAttr);
                    string sFeildName = oField.GetNameRef();

                    #region 獲取屬性字段  
                    FieldType Ftype = oFeature.GetFieldType(sFeildName);
                    switch (Ftype)
                    {
                        case FieldType.OFTString:
                            string sFValue = oFeature.GetFieldAsString(sFeildName);
                            string sTempType = "string";
                            FeildStringList.Add(sFValue);
                            break;
                        case FieldType.OFTReal:
                            double dFValue = oFeature.GetFieldAsDouble(sFeildName);
                            sTempType = "float";
                            FeildStringList.Add(dFValue.ToString());
                            break;
                        case FieldType.OFTInteger:
                            int iFValue = oFeature.GetFieldAsInteger(sFeildName);
                            sTempType = "int";
                            FeildStringList.Add(iFValue.ToString());
                            break;
                        default:
                            //sFValue = oFeature.GetFieldAsString(ChosenFeildIndex[iFeildIndex]);  
                            sTempType = "string";
                            break;
                    }
                    #endregion
                }
            }
            return true;
        }
        /// <summary>  
        /// 獲取數據  
        /// </summary>  
        /// <returns></returns>  
        public bool GetGeometry(int iIndex)
        {
            if (null == oLayer)
            {
                return false;
            }
            int iFeatureCout = (int)oLayer.GetFeatureCount(0);
            Feature oFeature = null;
            oFeature = oLayer.GetFeature(iIndex);
            //  Geometry  
            Geometry oGeometry = oFeature.GetGeometryRef();
            wkbGeometryType oGeometryType = oGeometry.GetGeometryType();
            switch (oGeometryType)
            {
                case wkbGeometryType.wkbPoint:
                    oGeometry.ExportToWkt(out sCoordiantes);
                    sCoordiantes = sCoordiantes.ToUpper().Replace("POINT (", "").Replace(")", "");
                    break;
                case wkbGeometryType.wkbLineString:
                case wkbGeometryType.wkbLinearRing:
                    oGeometry.ExportToWkt(out sCoordiantes);
                    sCoordiantes = sCoordiantes.ToUpper().Replace("LINESTRING (", "").Replace(")", "");
                    break;
                default:
                    break;
            }
            return false;
        }

    }//END class  
}