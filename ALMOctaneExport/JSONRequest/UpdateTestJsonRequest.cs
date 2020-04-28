using System;
using System.Globalization;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace ALMOctaneExport.JSONRequest
{
    public partial class UpdateJson
    {
        [JsonProperty("script")]
        public string Script { get; set; }

        [JsonProperty("comment")]
        public string Comment { get; set; }

        [JsonProperty("revision_type")]
        public string RevisionType { get; set; }
    }
}
