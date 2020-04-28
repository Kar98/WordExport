using MicroFocus.Adm.Octane.Api.Core.Entities;
using MicroFocus.Adm.Octane.Api.Core.Services;
using MicroFocus.Adm.Octane.Api.Core.Services.Query;
using MicroFocus.Adm.Octane.Api.Core.Services.RequestContext;
using System;
using System.Collections.Generic;

namespace ALMOctaneExport
{
    public static class TestHelper
    {
        public static string SUBTYPE_MANUAL_TEST = "test_manual";

        public static Phase GetPhaseForEntityByLogicalName(EntityService entityService, WorkspaceContext workspaceContext, String entityTypeName, String logicalName)
        {
            List<QueryPhrase> queryPhrases = new List<QueryPhrase>();
            LogicalQueryPhrase byEntityPhrase = new LogicalQueryPhrase(Phase.ENTITY_FIELD, entityTypeName);
            LogicalQueryPhrase byNamePhrase = new LogicalQueryPhrase(Phase.LOGICAL_NAME_FIELD, logicalName);
            queryPhrases.Add(byEntityPhrase);
            queryPhrases.Add(byNamePhrase);
            EntityListResult<Phase> result = entityService.Get<Phase>(workspaceContext, queryPhrases, null);
            Phase phase = result.data[0];
            return phase;
        }

        public static WorkItemRoot GetWorkItemRoot(EntityService entityService, WorkspaceContext workspaceContext)
        {
            List<String> fields = new List<String>() { Phase.NAME_FIELD };
            EntityListResult<WorkItemRoot> result = entityService.Get<WorkItemRoot>(workspaceContext, null, fields);
            
            WorkItemRoot root = result.data[0];
            return root;
        }

        public static ListNode GetSeverityByName(EntityService entityService, WorkspaceContext workspaceContext, String name)
        {
            String suffix = name.ToLower().Replace(" ", "_");
            String logicalName = "list_node.severity." + suffix;
            List<QueryPhrase> queryPhrases = new List<QueryPhrase>();
            LogicalQueryPhrase byLogicalName = new LogicalQueryPhrase(ListNode.LOGICAL_NAME_FIELD, logicalName);
            queryPhrases.Add(byLogicalName);

            // LogicalQueryPhrase byName = new LogicalQueryPhrase(ListNode.NAME_FIELD, name);
            //queryPhrases.Add(byName);

            List<String> fields = new List<String>() { Phase.NAME_FIELD, Phase.LOGICAL_NAME_FIELD };

            EntityListResult<ListNode> result = entityService.Get<ListNode>(workspaceContext, queryPhrases, fields);
            
            ListNode listNode = result.data[0];
            return listNode;
        }

        public static string[] NameFields = new string[] { "name" };
        public static string[] NameSubtypeFields = new string[] { "name", "subtype" };
    }
}
