namespace NTVXApi.Models
{
    public class ClientVersionModel
    {
        public string CurrentVersion { get; set; }
        public string NewVersion { get; set; }
        public int VersionNum { get; set; }
        public bool EndOfLife { get; set; }
    }
}