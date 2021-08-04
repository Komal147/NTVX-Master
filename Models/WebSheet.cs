using System.Collections.Generic;

namespace NTVXApi.Models
{
    public class WebSheet
    {
        public string Name { get; set; }
        public List<WebCell> Cells { get; set; }
    }
}