using NTVXApi.Models;
using System.Collections.Generic;
using System.Linq;

namespace NTVXApi.Services
{
    public static class ClientVersion
    {
        public static Dictionary<string, List<ClientVersionModel>> VersionMapping { get; } = new Dictionary<string, List<ClientVersionModel>>();

        public static MessageModel GetLatestVersion(string clientFile)
        {
            if (VersionMapping.ContainsKey(clientFile))
            {
                List<ClientVersionModel> versionList = VersionMapping[clientFile];
                ClientVersionModel latestVersion;

                //checking for end of life
                latestVersion = versionList.Where(m => m.EndOfLife == true).OrderByDescending(m => m.NewVersion).FirstOrDefault();
                if (latestVersion != null)
                {
                    return new MessageModel()
                    {
                        Message = Messages.EndOfLifeClient,
                        Model = latestVersion.NewVersion,
                        MessageType = MessageType.Critical
                    };
                }

                latestVersion = versionList.OrderByDescending(m => m.NewVersion).First();

                return new MessageModel()
                {
                    Message = "",
                    Model = latestVersion.NewVersion,
                    MessageType = MessageType.Success
                };
            }
            else
            {
                return new MessageModel()
                {
                    Message = Messages.InvalidClient,
                    Model = "",
                    MessageType = MessageType.Critical
                };
            }
        }
    }
}
