using NTVXApi.Models;
using NTVXApi.Services;
using System;
using System.Configuration;
using System.Globalization;
using System.Net;
using System.Web.Mvc;

namespace NTVXApi.Controllers
{
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Security", "CA5363:Do Not Disable Request Validation", Justification = "<Pending>")]
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Security", "CA3147:Mark Verb Handlers With Validate Antiforgery Token", Justification = "<Pending>")]
    public class ExcelController : Controller
    {
        AuthService auth;

        public ExcelController()
        {
            auth = new AuthService();
        }

        [ValidateInput(false)]
        [HttpPost]
        public ActionResult Process(NtvxModel input)
        {
            return new ExcelService(auth).Process(input);
        }

        [HttpPost]
        public ActionResult Login(string clientFile)
        {
            ActionResult result = new HttpStatusCodeResult(401);

            if(auth.Authenticate(true) == HttpStatusCode.OK)
            {
                if (auth.ByPassOtp)
                {
                    result = SessionStart(clientFile);
                }
                else
                {
                    string otp = auth.GenerateOTP();
                    auth.EmailOtp(otp);
                    result = Content("SUCCESS");
                }
            }

            return result;
        }

        [HttpPost]
        public ActionResult ValidateOtp(string otp, string clientFile)
        {
            if (!auth.ValidateOTP(otp))
            {
                throw new ApplicationException(Messages.WrongOTP);
            }

            return SessionStart(clientFile);
        }

        [HttpPost]
        public ActionResult SessionStart(string clientFile)
        {
            if (auth.Authenticate(false) == HttpStatusCode.Unauthorized)
            {
                return new HttpStatusCodeResult(401);
            }
            Telemetry.LogSessionStart(auth.UserNumber, clientFile);
            return Content(ClientVersion.GetLatestVersion(clientFile).ToString());
        }

        [HttpPost]
        public ActionResult SessionEnd()
        {
            Telemetry.LogSessionEnd();
            return Content("");
        }

        [HttpPost]
        public ActionResult GetClientFile(string clientFile)
        {
            string excelPath = Server.MapPath("~/Excel/ClientVersions/");
            excelPath = excelPath + clientFile;
            return File(excelPath, "application/vnd.ms-excel");
        }
    }
}
