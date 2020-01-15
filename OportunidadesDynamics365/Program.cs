using BIT.Entities;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Reflection;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;

namespace OportunidadesDynamics365
{
    public class Program
    {
        public static readonly string URI = System.Configuration.ConfigurationManager.AppSettings["URI"];
        public static readonly string USER = System.Configuration.ConfigurationManager.AppSettings["USER"];
        public static readonly string PASS = System.Configuration.ConfigurationManager.AppSettings["PASS"];
        public static readonly string DEVICE = System.Configuration.ConfigurationManager.AppSettings["DEVICE"];
        public static readonly string DEVICE_PASS = System.Configuration.ConfigurationManager.AppSettings["DEVICE_PASS"];

        private static string m_className = "[OportunidadesDynamics365.Program]";
        static void Main(string[] args)
        {
            string methodName = "[static void Main(string[] args)]";
            IOrganizationService _service = null;

            OrganizationServiceProxy _serviceProxy = null;

            try
            {
                var uri = new Uri(URI);
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                IServiceConfiguration<IOrganizationService> orgConfigInfo =
                    ServiceConfigurationFactory.CreateConfiguration<IOrganizationService>(uri);
                var creds = new ClientCredentials();
                ClientCredentials cred = new ClientCredentials();
                // set default credentials for OrganizationService
                cred.UserName.UserName = USER;
                cred.UserName.Password = PASS;

                ClientCredentials dev = new ClientCredentials();
                // set default credentials for OrganizationService
                dev.UserName.UserName = DEVICE;
                dev.UserName.Password = DEVICE_PASS;

                _serviceProxy = new OrganizationServiceProxy(uri, null, cred, dev);
                _serviceProxy.ServiceConfiguration.CurrentServiceEndpoint.Behaviors.Add(new ProxyTypesBehavior(Assembly.GetExecutingAssembly()));
                _serviceProxy.EnableProxyTypes();
                _service = (IOrganizationService)_serviceProxy;

                

                BITContext ctx = new BITContext(_serviceProxy);

                Account cliente = new Account(); //Early bound entity
                SystemUser comercial = new SystemUser();
                bit_modelo_comercial modelo = new bit_modelo_comercial();
                bit_unidaddenegocio unidad = new bit_unidaddenegocio();

                //Buscar el cliente
                cliente = ctx.AccountSet.Where(ac => ac.Name.Equals("ACRONIS")).FirstOrDefault();

                //Buscar comercial
                comercial = ctx.SystemUserSet.Where(cm => cm.DomainName.Equals("admin-crm@grupobusiness.it")).FirstOrDefault();

                //Buscar modelo comercial
                modelo = ctx.bit_modelo_comercialSet.Where(mc => mc.bit_name.Equals("SERVICIO ADMINISTRADO")).FirstOrDefault();

                //Buscar unidad negocio
                unidad = ctx.bit_unidaddenegocioSet.Where(u => u.bit_name.Equals("GESTION DOCUMENTAL")).FirstOrDefault();





                Opportunity op = new Opportunity();

                op.Name = "Oportunidad 15012020";
                op.ParentAccountId = new EntityReference(Account.EntityLogicalName,cliente.Id);
                op.OwnerId = new EntityReference(SystemUser.EntityLogicalName, comercial.Id);
                op.EstimatedCloseDate = DateTime.Today;
                op.OpportunityRatingCode = new OptionSetValue(1); // ALTA 1 MEDIA 2 BAJA 3
                op.bit_cuantos_productos = 1;
                op.bit_modelo_comercial = new EntityReference(bit_modelo_comercial.EntityLogicalName, modelo.Id);
                op.new_UnidadNegocio = new EntityReference(bit_unidaddenegocio.EntityLogicalName, unidad.Id);

                var response = _service.Create(op);

                //Consultar las categorias creadas

                //Crear el detalle del pago

                
                //Cambia la fase al 40%
                ChangeStage(_service, response);



                Console.WriteLine(response);

            }
            catch (MsgException ex)
            {
                ex.CallStack.Push(m_className, methodName);
                ex.Log(true, false);
            }
            catch (Exception ex)
            {
                MsgException oEx = new MsgException("", m_className, methodName, ex);
                oEx.Log(true, false);
            }



        }

        public static void ChangeStage(IOrganizationService service, Guid op)
        {
            // Get Process Instances
            RetrieveProcessInstancesRequest processInstanceRequest = new RetrieveProcessInstancesRequest
            {
                EntityId = op,
                EntityLogicalName = Opportunity.EntityLogicalName
            };

            RetrieveProcessInstancesResponse processInstanceResponse = (RetrieveProcessInstancesResponse)service.Execute(processInstanceRequest);

            // Declare variables to store values returned in response
            int processCount = processInstanceResponse.Processes.Entities.Count;
            Entity activeProcessInstance = processInstanceResponse.Processes.Entities[0]; // First record is the active process instance
            Guid activeProcessInstanceID = activeProcessInstance.Id; // Id of the active process instance, which will be used later to retrieve the active path of the process instance

            // Retrieve the active stage ID of in the active process instance
            Guid activeStageID = new Guid(activeProcessInstance.Attributes["processstageid"].ToString());

            // Retrieve the process stages in the active path of the current process instance
            RetrieveActivePathRequest pathReq = new RetrieveActivePathRequest
            {
                ProcessInstanceId = activeProcessInstanceID
            };
            RetrieveActivePathResponse pathResp = (RetrieveActivePathResponse)service.Execute(pathReq);

            string activeStageName = "";
            int activeStagePosition = -1;

            Console.WriteLine("\nRetrieved stages in the active path of the process instance:");
            for (int i = 0; i < pathResp.ProcessStages.Entities.Count; i++)
            {
                // Retrieve the active stage name and active stage position based on the activeStageId for the process instance
                if (pathResp.ProcessStages.Entities[i].Attributes["processstageid"].ToString() == activeStageID.ToString())
                {
                    activeStageName = pathResp.ProcessStages.Entities[i].Attributes["stagename"].ToString();
                    activeStagePosition = i;
                }
            }

            // Retrieve the stage ID of the next stage that you want to set as active
            activeStageID = (Guid)pathResp.ProcessStages.Entities[3].Attributes["processstageid"];

            // Retrieve the process instance record to update its active stage
            ColumnSet cols1 = new ColumnSet();
            cols1.AddColumn("activestageid");
            Entity retrievedProcessInstance = service.Retrieve("opportunitysalesprocess", activeProcessInstanceID, cols1);

            // Set the next stage as the active stage
            retrievedProcessInstance["activestageid"] = new EntityReference(OpportunitySalesProcess.EntityLogicalName, activeStageID);
            service.Update(retrievedProcessInstance);
        }

    }
}
