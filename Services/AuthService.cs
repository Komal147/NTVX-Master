using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Net;
using System.Net.Mail;
using System.Web;

namespace NTVXApi.Services
{
    public class AuthService
    {
        int MaxOTPByPass = 3;
        DateTime MaxOTPLastLogin = DateTime.Now.AddDays(-7);

        private int userNumber;
        private string email;
        private string fullName;
        private string product;
        private bool twoFactor;
        private bool isAuthenticated = false;
        private DateTime LastNTVXLogin;
        private int LastOTPCounter;

        public int UserNumber => userNumber;
        public bool ByPassOtp { get; set; }

        public HttpStatusCode Authenticate(bool twoFactor)
        {
            if (isAuthenticated)
            {
                return HttpStatusCode.OK;
            }

            // setting the two factor authentication
            this.twoFactor = twoFactor;

            // checking for the auth headers. If not present pushing the Auth headers. This will prompt username/password dialog in excel
            string auth = HttpContext.Current.Request.Headers["Authorization"];
            if (string.IsNullOrEmpty(auth))
            {
                return Unauthorized();
            }

            // validating username and password that is sent with Base64 encoding
            string authCred = auth.Substring(6);
            byte[] data = System.Convert.FromBase64String(authCred);
            string emailpassword = System.Text.ASCIIEncoding.ASCII.GetString(data);
            string[] usercreds = emailpassword.Split(':');
            isAuthenticated = false;

            string connectionString = ConfigurationManager.ConnectionStrings["BvxDB"].ConnectionString;
            using (SqlConnection cnn = new SqlConnection(connectionString))
            {
                string sql = @"SELECT nUserNo, vFirstName + ' ' + vLastName AS FullName, vProduct, LastNTVXLogin, LastOTPCounter
                                FROM UserMst u
                                WHERE lower(vEmail) = lower(@email)
                                  AND vField3 = @password
                                  AND EXISTS (SELECT 1 FROM PermissionMst p WHERE p.nUserNo = u.nUserNo)";
                cnn.Open();
                using (SqlCommand cmd = new SqlCommand(sql, cnn))
                {
                    cmd.Parameters.AddWithValue("@email", usercreds[0]);
                    cmd.Parameters.AddWithValue("@password", usercreds[1]);
                    SqlDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);

                    if (reader.Read())
                    {
                        email = usercreds[0];
                        userNumber = Convert.ToInt32(reader["nUserNo"], CultureInfo.InvariantCulture);
                        fullName = reader["FullName"].ToString();
                        product = reader["vProduct"] == DBNull.Value ? "" : reader["vProduct"].ToString();
                        LastNTVXLogin = reader["LastNTVXLogin"] == DBNull.Value ? MaxOTPLastLogin : Convert.ToDateTime(reader["LastNTVXLogin"], CultureInfo.InvariantCulture);
                        LastOTPCounter = reader["LastOTPCounter"] == DBNull.Value ? MaxOTPByPass : Convert.ToInt32(reader["LastOTPCounter"], CultureInfo.InvariantCulture);
                        if (LastOTPCounter >= MaxOTPByPass || LastNTVXLogin <= MaxOTPLastLogin)
                        {
                            ByPassOtp = false;
                        }
                        else
                        {
                            ByPassOtp = true;
                            UpdateLastOtpLogin(++LastOTPCounter);
                        }
                        isAuthenticated = true;
                        Telemetry.SetSessionID();
                    }
                    reader.Close();
                }
            }
            return AuthStatus(isAuthenticated);
        }

        public HttpStatusCode AuthorizeTemplateMacro(string templateFile, string templateMacro)
        {
            HttpStatusCode statusCode = Authenticate(false);
            if (statusCode == HttpStatusCode.Unauthorized)
            {
                return statusCode;
            }

            string connectionString = ConfigurationManager.ConnectionStrings["BvxDB"].ConnectionString;
            bool isAuthenticated = false;

            using (SqlConnection cnn = new SqlConnection(connectionString))
            {
                string sql = @"SELECT vMacroName FROM PermissionMst
                                WHERE nUserNo = @userNumber
                                  AND vExcelTemplate = @template";
                cnn.Open();
                using (SqlCommand cmd = new SqlCommand(sql, cnn))
                {
                    cmd.Parameters.AddWithValue("@userNumber", userNumber);
                    cmd.Parameters.AddWithValue("@template", templateFile);

                    SqlDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);

                    while (reader.Read())
                    {
                        if (reader[0] == DBNull.Value)
                        {
                            if (string.IsNullOrEmpty(templateMacro))
                                isAuthenticated = true;
                            continue;
                        }

                        string[] dbMacros = reader.GetString(0).Split(',');
                        foreach (string dbMacro in dbMacros)
                        {
                            if (dbMacro == "*" || dbMacro == templateMacro)
                            {
                                isAuthenticated = true;
                                break;
                            }
                        }
                    }

                    reader.Close();
                }
            }

            return AuthStatus(isAuthenticated);
        }

        private HttpStatusCode Unauthorized()
        {
            if (!this.twoFactor)
            {
                //HttpContext.Current.Response.Headers.Add("WWW-Authenticate", @"basic realm=""NTV Web API""");
            }
            return HttpStatusCode.Unauthorized;
        }

        private HttpStatusCode AuthStatus(bool isAuthenticated)
        {
            if (isAuthenticated)
            {
                return HttpStatusCode.OK;
            }
            else
            {
                return Unauthorized();
            }
        }

        public string GenerateOTP()
        {
            string otp = new Random().Next(100000, 1000000).ToString("D6", CultureInfo.InvariantCulture);
            //otp = "123456";
            HttpContext.Current.Session[UserNumber.ToString(CultureInfo.InvariantCulture)] = otp;
            return otp;
        }

        public bool ValidateOTP(string otp)
        {
            if (Authenticate(false) == HttpStatusCode.OK && HttpContext.Current.Session[UserNumber.ToString(CultureInfo.InvariantCulture)].ToString() == otp)
            {
                UpdateLastOtpLogin(0);
                return true;
            }
            else
            {
                return false;
            }
        }

        public void EmailOtp(string otp)
        {
            using (SmtpClient smtp = new SmtpClient())
            {
                using (MailMessage message = new MailMessage())
                {
                    message.To.Add(email);
                    message.Subject = $@"Your {product} log-in code";
                    message.Body = GetMailBody(otp);
                    message.IsBodyHtml = true;

                    smtp.Send(message);
                }
            }
        }

        private string GetMailBody(string otp)
        {
            return $@"Dear {fullName}, <br/><br/>
            This is an automated message for your {product} log-in code. <br/>
            Your code is<b>: {otp} </b><br>
            This code is one-time use only. Please do not send emails back to this unmonitored account";
        }

        private void UpdateLastOtpLogin(int otpCounter)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["BvxDB"].ConnectionString;
            using (SqlConnection cnn = new SqlConnection(connectionString))
            {
                string sql = @"UPDATE UserMst SET LastNTVXLogin = @LastNTVXLogin, 
                                LastOTPCounter = @LastOTPCounter
                                WHERE nUserNo = @userNumber";
                cnn.Open();
                using (SqlCommand cmd = new SqlCommand(sql, cnn))
                {
                    cmd.Parameters.AddWithValue("@LastNTVXLogin", DateTime.Now);
                    cmd.Parameters.AddWithValue("@LastOTPCounter", otpCounter);
                    cmd.Parameters.AddWithValue("@userNumber", userNumber);

                    cmd.ExecuteNonQuery();
                }
            }
        }
    }
}
