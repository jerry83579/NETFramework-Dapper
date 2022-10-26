using Dapper;
using FileTest.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace FileTest.Helpers
{
    public class SqlConn
    {
        public SqlConn()
        { }

        /// <summary>
        /// 取得資料
        /// </summary>
        /// <param name="conStr">連線字串</param>
        /// <returns></returns>
        public static List<dynamic> DbQuery(string conStr)
        {
            try
            {
                using (SqlConnection myConnection = new SqlConnection(ConfigurationManager.
                    ConnectionStrings["DefaultConnection"].ConnectionString))
                {
                    myConnection.Open();
                    return myConnection.Query(conStr).ToList();
                }
            }
            catch (SqlException exp)
            {
                throw new InvalidOperationException("Data could not be read", exp);
            }
        }

        public static int DbExecute(string conStr, List<Info> insertData)
        {
            try
            {
                using (SqlConnection myConnection = new SqlConnection(ConfigurationManager.
                    ConnectionStrings["DefaultConnection"].ConnectionString))
                {
                    myConnection.Open();
                    return myConnection.Execute(conStr, insertData);
                }
            }
            catch (SqlException exp)
            {
                throw new InvalidOperationException("Data could not be read", exp);
            }
        }
    }
}