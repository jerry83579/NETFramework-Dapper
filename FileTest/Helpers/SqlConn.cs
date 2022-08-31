﻿using Dapper;
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
        private SqlConnection myConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString);
      
        /// <summary>
        /// 取得資料
        /// </summary>
        /// <param name="conStr">連線字串</param>
        /// <returns></returns>
        public List<object> DbQuery (string conStr){
            try
            {
                myConnection.Open();
                var result = myConnection.Query(conStr).ToList();
                myConnection.Close();
                return result;
            }
            catch (SqlException exp)
            {
                throw new InvalidOperationException("Data could not be read", exp);
            }
        }
    }
}