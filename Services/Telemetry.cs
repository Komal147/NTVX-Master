using System;
using System.Configuration;
using System.Data.SqlClient;
using System.Web;

namespace NTVXApi.Services
{
    public static class Telemetry
    {
        public static string SessionID => HttpContext.Current.Request.Headers.Get("SessionID") ?? 
                                            HttpContext.Current.Response.Headers.Get("SessionID");

        public static void SetSessionID()
        {
            string sessionID = HttpContext.Current.Request.Headers.Get("SessionID");
            if (string.IsNullOrEmpty(sessionID))
            {
                sessionID = Guid.NewGuid().ToString();
            }
            HttpContext.Current.Response.Headers.Add("SessionID", sessionID);
        }

        public static void Log(int userNumber, string clientFile, string clientButton, 
            DateTime? startTime = null, DateTime? endTime = null, TimeSpan? duration = null)
        {
            // OnlienNTVTelemetry
            string connectionString = ConfigurationManager.ConnectionStrings["BvxDB"].ConnectionString;
            using (SqlConnection cnn = new SqlConnection(connectionString))
            {
                string sql = @"INSERT INTO OnlienNTVTelemetry(nUserNo, dStartTime, dToTime, tUsageDuration, 
                                dCreatedDate, dModifyDate, ClientFile, ClientButton, SessionID) 
                                VALUES(@nUserNo, @dStartTime, @dToTime, @tUsageDuration, 
                                @dCreatedDate, @dModifyDate, @ClientFile, @ClientButton, @SessionID)";
                cnn.Open();
                using (SqlCommand cmd = new SqlCommand(sql, cnn))
                {
                    cmd.Parameters.AddWithValue("@nUserNo", userNumber);
                    cmd.Parameters.AddWithValue("@dStartTime", startTime ?? DateTime.Now);
                    cmd.Parameters.AddWithValue("@dToTime", GetDbValue(endTime));
                    cmd.Parameters.AddWithValue("@tUsageDuration", GetDbValue(duration));
                    cmd.Parameters.AddWithValue("@dCreatedDate", DateTime.Now);
                    cmd.Parameters.AddWithValue("@dModifyDate", DateTime.Now);
                    cmd.Parameters.AddWithValue("@ClientFile", GetDbValue(clientFile));
                    cmd.Parameters.AddWithValue("@ClientButton", GetDbValue(clientButton));
                    cmd.Parameters.AddWithValue("@SessionID", Telemetry.SessionID);

                    cmd.ExecuteNonQuery();
                }
            }
        }

        public static void LogSessionStart(int userNumber, string clientFile)
        {
            Telemetry.Log(userNumber, clientFile, "");
        }

        public static void LogInteraction(int userNumber, string clientFile, string clientButton, DateTime startTime, DateTime endTime, TimeSpan duration)
        {
            Telemetry.Log(userNumber, clientFile, clientButton, startTime, endTime, duration);
        }

        public static void LogSessionEnd()
        {
            if (string.IsNullOrEmpty(Telemetry.SessionID))
                return;

            string connectionString = ConfigurationManager.ConnectionStrings["BvxDB"].ConnectionString;
            using (SqlConnection cnn = new SqlConnection(connectionString))
            {
                DateTime startTime = DateTime.Now;
                int telemetryId = 0;

                // getting the previously inserted Session start
                string sql = @"SELECT TOP 1 nOnlienNTVTelemetry, dStartTime FROM OnlienNTVTelemetry 
                        WHERE dToTime IS NULL AND SessionID = @SessionID ORDER BY nOnlienNTVTelemetry";
                cnn.Open();
                using (SqlCommand cmdQuery = new SqlCommand(sql, cnn))
                {
                    cmdQuery.Parameters.AddWithValue("@SessionID", Telemetry.SessionID);
                    SqlDataReader reader = cmdQuery.ExecuteReader();

                    if (reader.Read())
                    {
                        telemetryId = (int) reader.GetDecimal(0);
                        startTime = reader.GetDateTime(1);
                    }
                    reader.Close();
                }

                // setting the endtime & timespan
                TimeSpan usageDuration = DateTime.Now - startTime;
                TimeSpan minDBDuration = TimeSpan.Zero;
                TimeSpan maxDBDuration = TimeSpan.FromDays(1) - TimeSpan.FromSeconds(1);

                if(usageDuration < minDBDuration)
                {
                    usageDuration = minDBDuration;
                }
                else if(usageDuration > maxDBDuration)
                {
                    usageDuration = maxDBDuration;
                }

                sql = @"UPDATE OnlienNTVTelemetry SET dToTime = @dToTime, dModifyDate = @dModifyDate, tUsageDuration = @tUsageDuration 
                        WHERE nOnlienNTVTelemetry = @TelemetryID";
                using (SqlCommand cmdUpdate = new SqlCommand(sql, cnn))
                {
                    cmdUpdate.Parameters.AddWithValue("@TelemetryID", telemetryId);
                    cmdUpdate.Parameters.AddWithValue("@dToTime", DateTime.Now);
                    cmdUpdate.Parameters.AddWithValue("@dModifyDate", DateTime.Now);
                    cmdUpdate.Parameters.AddWithValue("@tUsageDuration", usageDuration);
                    cmdUpdate.ExecuteNonQuery();
                }
            }
        }

        private static object GetDbValue(string value)
        {
            return string.IsNullOrEmpty(value) ? DBNull.Value : (object)value;
        }

        private static object GetDbValue(DateTime? value)
        {
            return (object) value ?? DBNull.Value;
        }

        private static object GetDbValue(TimeSpan? value)
        {
            return (object)value ?? DBNull.Value;
        }
    }
}