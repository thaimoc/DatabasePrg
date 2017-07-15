using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Core;
using System.Data.SqlClient;
using System.Data;

namespace DataAccess
{
    public static class Login
    {
        //string username;

        //public string Username
        //{
        //    get { return username; }
        //    set { username = value; }
        //}
        //string password;

        //public string Password
        //{
        //    get { return password; }
        //    set { password = value; }
        //}
        public static int UserLogin(string user, string pass)
        {
            string connectionString = "Server=PING-PC;UID=sa;PWD=Nguyen;Database=Shopes";
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand("User_Login2", con);
                cmd.CommandType = CommandType.StoredProcedure;
                con.Open();
                cmd.Parameters.Add("@User", SqlDbType.NVarChar).Value = user;
                cmd.Parameters.Add("@Pass", SqlDbType.NVarChar).Value = pass;
                using (SqlDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                {
                    while (reader.Read())
                    {
                        if (reader.GetString(1) == user && reader.GetString(2) == pass)
                            return 1;
                    }
                    return 0;
                }
            }
        }
    }
}
