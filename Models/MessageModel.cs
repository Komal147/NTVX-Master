namespace NTVXApi.Models
{
    public class MessageModel
    {
        public string Message { get; set; }
        public string Model { get; set; }
        public MessageType MessageType { get; set; }

        public override string ToString()
        {
            return Message + "~" + Model + "~" + (int)MessageType;
        }
    }
}
