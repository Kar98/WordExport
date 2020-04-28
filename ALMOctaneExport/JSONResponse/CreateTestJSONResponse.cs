using System;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace ALMOctaneExport.JSONResponse
{
    public partial class CreateTestJsonResponse
    {
        [JsonProperty("total_count")]
        public long TotalCount { get; set; }

        [JsonProperty("data")]
        public Datum[] Data { get; set; }

        [JsonProperty("exceeds_total_count")]
        public bool ExceedsTotalCount { get; set; }

        [JsonProperty("total_error_count")]
        public long TotalErrorCount { get; set; }
    }

    public partial class Datum
    {
        [JsonProperty("type")]
        public string Type { get; set; }

        [JsonProperty("covered_content")]
        public CoveredContent CoveredContent { get; set; }

        [JsonProperty("covered_manual_test")]
        public object CoveredManualTest { get; set; }

        [JsonProperty("covered_requirement")]
        public CoveredContent CoveredRequirement { get; set; }

        [JsonProperty("covering_automated_test")]
        public object CoveringAutomatedTest { get; set; }

        [JsonProperty("framework")]
        public object Framework { get; set; }

        [JsonProperty("modified_by")]
        public object ModifiedBy { get; set; }

        [JsonProperty("owner")]
        public object Owner { get; set; }

        [JsonProperty("product_areas")]
        public CoveredContent ProductAreas { get; set; }

        [JsonProperty("quality_stories")]
        public CoveredContent QualityStories { get; set; }

        [JsonProperty("scm_repository")]
        public object ScmRepository { get; set; }

        [JsonProperty("test_data_table")]
        public object TestDataTable { get; set; }

        [JsonProperty("test_level")]
        public object TestLevel { get; set; }

        [JsonProperty("test_runner")]
        public object TestRunner { get; set; }

        [JsonProperty("user_tags")]
        public CoveredContent UserTags { get; set; }

        [JsonProperty("description")]
        public string Description { get; set; }

        [JsonProperty("steps_num")]
        public long StepsNum { get; set; }

        [JsonProperty("approved_version")]
        public object ApprovedVersion { get; set; }

        [JsonProperty("workspace_id")]
        public long WorkspaceId { get; set; }

        [JsonProperty("latest_version")]
        public long LatestVersion { get; set; }

        [JsonProperty("executable")]
        public bool Executable { get; set; }

        [JsonProperty("package")]
        public object Package { get; set; }

        [JsonProperty("version_stamp")]
        public long VersionStamp { get; set; }

        [JsonProperty("class_name")]
        public object ClassName { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("sha")]
        public object Sha { get; set; }

        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("last_modified")]
        public DateTimeOffset LastModified { get; set; }

        [JsonProperty("subtype")]
        public string Subtype { get; set; }

        [JsonProperty("component")]
        public object Component { get; set; }

        [JsonProperty("automation_identifier")]
        public object AutomationIdentifier { get; set; }

        [JsonProperty("created")]
        public DateTimeOffset Created { get; set; }

        [JsonProperty("identity_hash")]
        public object IdentityHash { get; set; }

        [JsonProperty("script_path")]
        public object ScriptPath { get; set; }

        [JsonProperty("estimated_duration")]
        public object EstimatedDuration { get; set; }

        [JsonProperty("manual")]
        public bool Manual { get; set; }

        [JsonProperty("creation_time")]
        public DateTimeOffset CreationTime { get; set; }

        [JsonProperty("is_draft")]
        public bool IsDraft { get; set; }

        [JsonProperty("client_lock_stamp")]
        public long ClientLockStamp { get; set; }

        [JsonProperty("num_comments")]
        public long NumComments { get; set; }

        [JsonProperty("pipelines")]
        public string Pipelines { get; set; }

        [JsonProperty("builds")]
        public string Builds { get; set; }

        [JsonProperty("test_status")]
        public string TestStatus { get; set; }

        [JsonProperty("has_comments")]
        public bool HasComments { get; set; }

        [JsonProperty("run_in_releases")]
        public string RunInReleases { get; set; }

        [JsonProperty("has_attachments")]
        public bool HasAttachments { get; set; }

        [JsonProperty("global_text_search_result")]
        public object GlobalTextSearchResult { get; set; }

        [JsonProperty("assigned_to_me")]
        public bool AssignedToMe { get; set; }

        [JsonProperty("followed_by_me")]
        public bool FollowedByMe { get; set; }

        [JsonProperty("subtype_label")]
        public string SubtypeLabel { get; set; }

        [JsonProperty("phase")]
        public AutomationStatus Phase { get; set; }

        [JsonProperty("testing_tool_type")]
        public AutomationStatus TestingToolType { get; set; }

        [JsonProperty("author")]
        public Author Author { get; set; }

        [JsonProperty("automation_status")]
        public AutomationStatus AutomationStatus { get; set; }

        [JsonProperty("designer")]
        public Author Designer { get; set; }

        [JsonProperty("test_type")]
        public CoveredContent TestType { get; set; }
    }

    public partial class Author
    {
        [JsonProperty("type")]
        public string Type { get; set; }

        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("workspace_id")]
        public long WorkspaceId { get; set; }

        [JsonProperty("activity_level")]
        public long ActivityLevel { get; set; }

        [JsonProperty("full_name")]
        public string FullName { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }
    }

    public partial class AutomationStatus
    {
        [JsonProperty("type")]
        public string Type { get; set; }

        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("logical_name")]
        public string LogicalName { get; set; }

        [JsonProperty("activity_level")]
        public long ActivityLevel { get; set; }

        [JsonProperty("index")]
        public long Index { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("workspace_id", NullValueHandling = NullValueHandling.Ignore)]
        public long? WorkspaceId { get; set; }
    }

    public partial class CoveredContent
    {
        [JsonProperty("total_count")]
        public long TotalCount { get; set; }

        [JsonProperty("data")]
        public AutomationStatus[] Data { get; set; }
    }
}
