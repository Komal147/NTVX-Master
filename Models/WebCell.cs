namespace NTVXApi.Models
{
    public class WebCell
    {
        public WebCell() { }
        public WebCell(string key, string value)
        {
            this.Key = key;
            this.Value = value;
        }

        public string Key { get; set; }
        public string Value { get; set; }
    }
}