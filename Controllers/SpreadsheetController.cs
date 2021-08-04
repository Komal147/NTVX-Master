using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using NTVXApi.Models;
using NTVXApi.Services;
using Syncfusion.EJ2.Spreadsheet;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Web.Mvc;
using System.Xml;

namespace NTVXApi.Controllers
{
    public class SpreadsheetController : Controller
    {

        [HttpPost]
        public ActionResult Open(OpenRequest openRequest)
        {
            return Content(Workbook.Open(openRequest));
        }

        [ValidateInput(false)]
        [HttpPost]
        public ActionResult Save(SaveSettings saveSettings)
        {
            Stream stream = Workbook.Save<Stream>(saveSettings);
            return File(stream, saveSettings.GetContentType(), saveSettings.FileName + "." + saveSettings.SaveType.ToString().ToLower());
        }

        [HttpGet]
        public ActionResult Template(int templateId)
        {
            WebTemplate template = WebTemplateCache.Templates.Find(t => t.TemplateId == templateId);
            return JsonContent(template);
        }

        [HttpGet]
        public ActionResult File(int templateId)
        {
            WebTemplate template = WebTemplateCache.Templates.Find(t => t.TemplateId == templateId);
            return File(Server.MapPath("~/Excel/WebTemplates/" + template.TemplateFile), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        }

        [ValidateInput(false)]
        [HttpPost]
        public ActionResult Process(WebModel input)
        {

            WebTemplate template = WebTemplateCache.Templates.Find(t => t.TemplateId == input.TemplateId);
            WebAction action = template.Actions.Find(a => a.Button == input.Button);

            string[] ranges = action.PasteSheetsAndRanges.Split(';');
            string returnSheetsAndRanges = "";
            
            foreach(string range in ranges)
            {
                string[] item = range.Split('~');
                returnSheetsAndRanges += string.IsNullOrEmpty(returnSheetsAndRanges) ? "" : ";";
                returnSheetsAndRanges += item[0] + "~" + item[1];
            }


            // processing input
            NtvxModel model = new NtvxModel()
            {
                ClientFile = "Web",
                ClientButton = input.Button,
                TemplateFile = template.ServerTemplate,
                InputXmlString = GetXml(input.Sheets),
                CopySheetsAndRanges = action.CopySheetsAndRanges,
                ReturnSheetsAndRanges = returnSheetsAndRanges,
                MacroToExecute = action.Macro,
                DeleteFile = template.DeleteFile
            };

            ExcelService excel = new ExcelService(new AuthService());
            ActionResult result = excel.Process(model);

            if (result is HttpStatusCodeResult)
            {
                HttpContext.Response.Headers.Add("WWW-Authenticate", @"basic realm=""NTV Web API""");
                return result;
            }

            // preparing output
            WebModel output = new WebModel()
            {
                TemplateId = template.TemplateId,
                Button = input.Button,
                PasteSheetsAndRange = action.PasteSheetsAndRanges,
                Sheets = new List<WebSheet>()
            };

            XmlDocument xmlDoc = new XmlDocument();
            string xml = ((ContentResult)result).Content;
            xmlDoc.LoadXml(xml);

            XmlNode root = xmlDoc.ChildNodes[0];

            WebSheet sheet;
            string sheetName;
            foreach (XmlNode xmlSheet in root.ChildNodes)
            {
                sheetName = xmlSheet.Attributes["name"].Value;
                sheet = output.Sheets.Find(s => s.Name == sheetName);
                if(sheet == null)
                {
                    sheet = new WebSheet() { Name = sheetName, Cells = new List<WebCell>() };
                    output.Sheets.Add(sheet);
                }

                foreach (XmlNode cell in xmlSheet.ChildNodes)
                {
                    sheet.Cells.Add(new WebCell()
                    {
                        Key = cell.Attributes["address"].Value,
                        Value = cell.InnerText
                    });
                }
            }

            return JsonContent(output);
        }

        private ActionResult JsonContent(object json)
        {
            var camelCaseFormatter = new JsonSerializerSettings();
            camelCaseFormatter.ContractResolver = new CamelCasePropertyNamesContractResolver();
            camelCaseFormatter.NullValueHandling = NullValueHandling.Ignore;
            return Content(JsonConvert.SerializeObject(json, camelCaseFormatter));
        }

        private string GetXml(List<WebSheet> sheets)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(@"<ntvx>");

            foreach(WebSheet sheet in sheets)
            {
                builder.Append(string.Format(@"<sheet name=""{0}"">", sheet.Name));

                foreach (WebCell cell in sheet.Cells)
                {
                    builder.Append(string.Format(@"<cell address= ""{0}"" >{1}</cell>", cell.Key, cell.Value));
                }

                builder.Append("</sheet>");
            }


            builder.Append("</ntvx>");
            return builder.ToString();
        }

        private void SetCorsHeaders()
        {
            var context = this.HttpContext;
            context.Response.Clear();
            string uri = context.Request.UrlReferrer.AbsoluteUri;
            uri = uri.Substring(0, uri.IndexOf('/', 8)); // take after http://
            context.Response.Headers.Add("Access-Control-Allow-Origin", uri);
            context.Response.Headers.Add("Access-Control-Allow-Methods", "POST");
            context.Response.Headers.Add("Access-Control-Allow-Credentials", "true");
        }
    }
}