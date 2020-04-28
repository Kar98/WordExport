using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ALMOctaneExport.JSONRequest;
using ALMOctaneExport.JSONResponse;
using MicroFocus.Adm.Octane.Api.Core.Connector;
using MicroFocus.Adm.Octane.Api.Core.Connector.Authentication;
using MicroFocus.Adm.Octane.Api.Core.Services;
using MicroFocus.Adm.Octane.Api.Core.Services.RequestContext;
using Newtonsoft.Json;

namespace ALMOctaneExport
{
    public class ALMOctaneAPI
    {
        public string UserId { get; set; }

        ALMOctaneConnection conn;
        public ALMOctaneAPI(ALMOctaneConnection conn)
        {
           this.conn = conn;
//            UserId = "16001";
        }

        /// <summary>
        /// Creates a new Manual Test and returns the ID
        /// </summary>
        /// <param name="name"></param>
        /// <param name="description"></param>
        /// <returns></returns>
        public string CreateTest(string name, string description)
        {
            string api = $"/api/shared_spaces/{conn.SharedspaceId}/workspaces/{conn.WorkspaceId}/tests";
            string apiParam = "fields=test_runner,creation_time,covered_content,version_stamp,script_path,covered_manual_test,workspace_id,num_comments,pipelines,builds,last_modified,approved_version,phase,test_status,subtype_label,client_lock_stamp,package,created,author,product_areas,estimated_duration,sha,user_tags,testing_tool_type,has_comments,automation_identifier,name,automation_status,scm_repository,covering_automated_test,run_in_releases,assigned_to_me,description,manual,followed_by_me,latest_version,steps_num,subtype,is_draft,class_name,owner,has_attachments,quality_stories,test_level,global_text_search_result,designer,covered_requirement,test_type,executable,identity_hash,component,framework,test_data_table,modified_by";
            var reqType = RequestType.Post;

            var obj = JsonConvert.DeserializeObject<CreateTestJsonRequest>(CreateTestJsonRequest.Default);

            obj.Data[0].Name = name;
            obj.Data[0].Description = description;
            obj.Data[0].Designer.Id = UserId;
            obj.Data[0].Author.Id = UserId;

            var json = JsonConvert.SerializeObject(obj);
            File.WriteAllText("json.txt", json);

            var res = conn.RestConnector.Send(api, apiParam, reqType, json);

            return JsonConvert.DeserializeObject<CreateTestJsonResponse>(res.Data).Data[0].Id;
        }

        public void UpdateTest(string testid, string script)
        {
            string api = $"/api/shared_spaces/{conn.SharedspaceId}/workspaces/{conn.WorkspaceId}/tests/{testid}/script";
            string apiParam = "fields=test_runner,creation_time,covered_content,version_stamp,script_path,covered_manual_test,workspace_id,num_comments,pipelines,builds,last_modified,approved_version,phase,test_status,subtype_label,client_lock_stamp,package,created,author,product_areas,estimated_duration,sha,user_tags,testing_tool_type,has_comments,automation_identifier,name,automation_status,scm_repository,covering_automated_test,run_in_releases,assigned_to_me,description,manual,followed_by_me,latest_version,steps_num,subtype,is_draft,class_name,owner,has_attachments,quality_stories,test_level,global_text_search_result,designer,covered_requirement,test_type,executable,identity_hash,component,framework,test_data_table,modified_by";
            var reqType = RequestType.Update;
            string data = @"{""script"": """",""comment"": """",""revision_type"": ""Minor""}";

            var obj = JsonConvert.DeserializeObject<UpdateJson>(data);
            obj.Script = script;

            var json = JsonConvert.SerializeObject(obj);

            conn.RestConnector.Send(api, apiParam, reqType, json);
        }
    }
}
