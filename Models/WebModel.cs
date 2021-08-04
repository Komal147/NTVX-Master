using System.Collections.Generic;

namespace NTVXApi.Models
{
    public class WebModel
    {
        public int TemplateId { get; set; }
        public string Button { get; set; }
        public string PasteSheetsAndRange { get; set; }
        public List<WebSheet> Sheets { get; set; }
    }
}
