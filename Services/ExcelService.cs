using NTVXApi.Models;
using System;
using System.Globalization;
using System.Net;
using System.Web;
using System.Web.Mvc;

namespace NTVXApi.Services
{
    public class ExcelService
    {
        AuthService auth;

        public ExcelService(AuthService auth)
        {
            this.auth = auth;
        }

        public ActionResult Process(NtvxModel input)
        {
            input.MacroToExecute = input.MacroToExecute ?? "";


            if (auth.AuthorizeTemplateMacro(input.TemplateFile, input.MacroToExecute) == HttpStatusCode.Unauthorized)
            {
                return new HttpStatusCodeResult(401);
            }

            MessageModel versionCheck = ClientVersion.GetLatestVersion(input.ClientFile);
            if (versionCheck.MessageType == MessageType.Critical)
            {
                throw new ApplicationException(versionCheck.Message);
            }

            // cloning the template file. The template file name is passed as input parameter
            string templatefile = HttpContext.Current.Server.MapPath("~/Excel/Templates/" + input.TemplateFile);

            if (!System.IO.File.Exists(templatefile))
            {
                throw new ApplicationException(Messages.InvalidTemplateFile);
            }

            string excelPath = HttpContext.Current.Server.MapPath("~/Excel");
            int extensionIndex = templatefile.LastIndexOf('.');
            string fileExtension = templatefile.Substring(extensionIndex, templatefile.Length - extensionIndex);
            string fileName = Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture) + fileExtension;
            string excelFile = excelPath + @"\" + fileName;

            DateTime startTime = DateTime.Now;
            try
            {
                System.IO.File.Copy(templatefile, excelFile);

                using (ExcelWorkbook excel = new ExcelWorkbook())
                {
                    // opening the template file (cloned) and calling the ServerProcess macro with all the input parameters
                    excel.Open(excelFile);
                    //excel.Open(templatefile);
                    string ret = excel.RunMacro("ServerModule.ServerProcess", input.InputXmlString, input.CopySheetsAndRanges, input.MacroToExecute, input.ReturnSheetsAndRanges, input.DownloadMacro, input.DownloadFileName);
                    return new ContentResult() { Content = ret };
                }
            }
            finally
            {
                DateTime endTime = DateTime.Now;
                TimeSpan duration = endTime - startTime;

                Telemetry.LogInteraction(auth.UserNumber, input.ClientFile, input.ClientButton, startTime, endTime, duration);
                try
                {
                    if (System.IO.File.Exists(excelFile) && input.DeleteFile)
                        System.IO.File.Delete(excelFile);
                }
#pragma warning disable CA1031 // Do not catch general exception types
                catch
#pragma warning restore CA1031 // Do not catch general exception types
                {
                    // ignore exception
                }
            }
        }
    }
}