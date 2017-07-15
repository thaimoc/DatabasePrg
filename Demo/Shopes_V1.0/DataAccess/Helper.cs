using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Configuration;
using Core;


namespace DataAccess
{
    public class Helper
    {
        public static string connectionString = "";

        public static SqlConnection GetSqlConnection() 
        {
            return new SqlConnection(connectionString);
        }
        public static Boolean TestConnect()
        {        
            Boolean result = false;
            using (SqlConnection connect = new SqlConnection(connectionString))
            {
                try
                {
                    connect.Open();
                    if (connect.State == System.Data.ConnectionState.Open)
                        result = true;
                }
                catch
                { 
                    
                }                
            }
            return result;
        }
    }
}
