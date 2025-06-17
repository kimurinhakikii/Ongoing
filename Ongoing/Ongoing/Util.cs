using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;

namespace Ongoing
{
     class Util
    {
        protected static string dataSource = @"10.128.223.72";
        protected static string dataBase = "PORTAL_CORPORATIVO_PRD";
        protected static string dbUserID = "USERDSC";        
        protected static string dbUserPWD = "v!V0__2O!8";
        protected static string connectionString()
        {
            return @"server=" + dataSource + "; user ID=" + dbUserID + ";password=" + dbUserPWD + ";Initial Catalog=" + dataBase + "; App=MAILLING " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
        }
        public static SqlConnection Conn()
        {
            try
            {
                SqlConnection conn = new SqlConnection();
                conn.ConnectionString = connectionString();
                conn.Open();
                return conn;
            }
            catch
            {
                return null;
            }
        }

        public static SqlCommand sqlReader(string sql)
        {
            try
            {
                SqlConnection conn = Conn();
                SqlCommand command = new SqlCommand(sql, conn);
                command.CommandTimeout = 0;

                return command;
            }
            catch
            {
                return null;
            }
        }

        public static bool sqlExecute(string sql)
        {
            try
            {
                SqlConnection conn = Conn();
                SqlCommand command = new SqlCommand(sql, conn);
                command.CommandTimeout = 0;
                command.ExecuteNonQuery();
                conn.Close();
                conn.Dispose();

                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
