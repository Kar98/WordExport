using System;
using System.Collections.Generic;

using System.Globalization;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace TestRailExport
{
    public partial class TestrailJSON
    {
        [JsonProperty("title")]
        public string Title { get; set; }

        [JsonProperty("template_id")]
        public long TemplateId { get; set; }

        [JsonProperty("type_id")]
        public long TypeId { get; set; }

        [JsonProperty("priority_id")]
        public long PriorityId { get; set; }

        [JsonProperty("custom_testscenario")]
        public string CustomTestscenario { get; set; }

        [JsonProperty("custom_steps_separated")]
        public CustomStepsSeparated[] CustomStepsSeparated { get; set; }
    }

    public partial class CustomStepsSeparated
    {
        [JsonProperty("content")]
        public string Content { get; set; }

        [JsonProperty("expected")]
        public string Expected { get; set; }
    }


}
