using Gurock.TestRail;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;

namespace TestRailExport
{
    // Create connection to TR
    // Create test case
    // Edit test case

    public class TRExporter
    {
        APIClient client;

        public string Token { get; private set; }

        public TRExporter(string host, string username, string password)
        {
            client = new APIClient(host);
            client.User = username;
            client.Password = password;
            Token = Convert.ToBase64String(
                Encoding.ASCII.GetBytes(
                    String.Format(
                        "{0}:{1}",
                        client.User,
                        client.Password
                    )
                )
            );
        }

        public long CreateTest(string sectionId, object bodyData)
        {
            string api = $"add_case/{sectionId}";
            return ((JObject)client.SendPost(api, bodyData)).Value<long>("id");
        }

        public void UpdateTest(string testcaseId, string bodyData)
        {
            string api = $"update_case/{testcaseId}";
            client.SendPost(api, bodyData);
        }
    }
}
