using System;
using System.Collections.Generic;

using System.Globalization;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace ALMOctaneExport.JSONRequest
{


    public partial class CreateTestJsonRequest
    {
        [JsonProperty("data")]
        public Datum[] Data { get; set; }

        public static string Default = @"{""data"":[{""designer"":{""type"":""workspace_user"",""id"":""16001""},""test_type"":{""data"":[{""type"":""list_node"",""id"":""list_node.test_type.end_to_end""}]},""subtype"":""test_manual"",""phase"":{""type"":""phase"",""id"":""phase.test_manual.new""},""author"":{""type"":""workspace_user"",""id"":""16001""},""description"":""<html><body><p>some description goes here</p>\n<p>extra line</p></body></html>"",""name"":""brandnew""}]}";
    }

    public partial class Datum
    {
        [JsonProperty("designer")]
        public Author Designer { get; set; }

        [JsonProperty("test_type")]
        public TestType TestType { get; set; }

        [JsonProperty("subtype")]
        public string Subtype { get; set; }

        [JsonProperty("phase")]
        public Author Phase { get; set; }

        [JsonProperty("author")]
        public Author Author { get; set; }

        [JsonProperty("description")]
        public string Description { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }
    }

    public partial class Author
    {
        [JsonProperty("type")]
        public string Type { get; set; }

        [JsonProperty("id")]
        public string Id { get; set; }
    }

    public partial class TestType
    {
        [JsonProperty("data")]
        public Author[] Data { get; set; }
    }
}
