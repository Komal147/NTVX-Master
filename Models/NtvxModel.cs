using System.Web.Mvc;

namespace NTVXApi.Models
{
    public class NtvxModel
    {
        public string ClientFile { get; set; }
        public string ClientButton { get; set; }
        public string TemplateFile { get; set; }
        [AllowHtml]
        public string InputXmlString { get; set; }
        public string CopySheetsAndRanges { get; set; }
        public string MacroToExecute { get; set; }
        public string ReturnSheetsAndRanges { get; set; }
        public bool DeleteFile { get; set; }

        // for dowloading the excel file
        public string DownloadMacro { get; set; }
        public string DownloadFileName { get; set; }
    }
}
