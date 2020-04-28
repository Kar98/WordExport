using MicroFocus.Adm.Octane.Api.Core.Connector;
using MicroFocus.Adm.Octane.Api.Core.Connector.Authentication;
using MicroFocus.Adm.Octane.Api.Core.Services;
using MicroFocus.Adm.Octane.Api.Core.Services.RequestContext;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace ALMOctaneExport
{
    public class ALMOctaneConnection : IDisposable
    {
        public RestConnector RestConnector { get; private set; }
        public EntityService EntityService { get; private set; }

        public string SharedspaceId { get; set; }
        public string WorkspaceId { get; set; }
        
        public ALMOctaneConnection(string host, string userName, string password, string sharedSpaceId, string workspaceId)
        {
            RestConnector = new RestConnector();
            EntityService = new EntityService(RestConnector);
            SharedspaceId = sharedSpaceId;
            WorkspaceId = workspaceId;

            LwssoAuthenticationStrategy lwssoAuthenticationStrategy;
            
            //WorkspaceContext workspaceContext;
            //SharedSpaceContext sharedSpaceContext;

            NetworkSettings.IgnoreServerCertificateValidation();
            NetworkSettings.EnableAllSecurityProtocols();

            ConnectionInfo connectionInfo;
            
            connectionInfo = new UserPassConnectionInfo(userName, password);
            lwssoAuthenticationStrategy = new LwssoAuthenticationStrategy(connectionInfo);
            RestConnector.Connect(host, lwssoAuthenticationStrategy);
            //workspaceContext = new WorkspaceContext(sharedSpaceId, workspaceId);
            //sharedSpaceContext = new SharedSpaceContext(sharedSpaceId);
        }

        public void Dispose()
        {
            RestConnector.Disconnect();
        }

    }
}
