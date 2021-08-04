using System.Collections.Generic;

namespace NTVXApi.Models
{
    public class WebTemplate
    {
        public int TemplateId { get; set; }
        public string TemplateFile { get; set; }
        public string ServerTemplate { get; set; }
        public List<WebAction> Actions { get; set; }
        public bool DeleteFile { get; set; }
    }
}
