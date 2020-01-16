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
                #region ConexionDyn365
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
                #endregion

                //SQL

                Money pvp = new Money(0), costo = new Money(0), costoIndirecto = new Money(0);

                BITContext ctx = new BITContext(_serviceProxy);
                Opportunity opportunidadCrm = new Opportunity();
                Account cliente = new Account(); //Early bound entity                                                 
                cliente = ctx.AccountSet.Where(ac => ac.new_RUC.Equals("1234567891")).FirstOrDefault();

                if (cliente == null)
                    throw new MsgException("No existe cliente con RUC 1234567891", m_className, methodName);

                opportunidadCrm = ctx.OpportunitySet.Where(opp => opp.Name.Equals("Servicios Cloud Azure - Icaro") && opp.StateCode == OpportunityState.Open && opp.ParentAccountId.Equals(cliente.Id)).FirstOrDefault();

                //Consultar categoria
                bit_categoria bit_CatalogoCategoria = ctx.bit_categoriaSet.Where(cat => cat.bit_categoria1.Equals("CONSUMO AZURE")).FirstOrDefault();

                if (opportunidadCrm != null)
                {
                    bit_pago bit_CategoriaOportunidad = ctx.bit_pagoSet.Where(cat => cat.bit_categoria_item.Equals(bit_CatalogoCategoria.Id)).FirstOrDefault();

                    if(bit_CategoriaOportunidad != null)
                    {

                        //Crear categoria 
                        Guid idCategoria = CrearCategoria(_service, opportunidadCrm.OwnerId.Id, opportunidadCrm.Id, bit_CatalogoCategoria.Id);

                        //Crear el item de la factura - crear por cada detalle de Icaro
                        CrearDetallePago(_service, opportunidadCrm.OwnerId.Id, opportunidadCrm.Id, bit_CatalogoCategoria.Id, bit_CategoriaOportunidad.Id, pvp, costo, costoIndirecto);

                        //Actualizar total oportunidad
                        ActualizarValoresOportunidad(_service, opportunidadCrm.Id, pvp, costo, costoIndirecto);
                    }
                    else
                    {

                        //Crear el item de la factura - crear por cada detalle de Icaro
                        CrearDetallePago(_service, opportunidadCrm.OwnerId.Id, opportunidadCrm.Id, bit_CatalogoCategoria.Id, bit_CategoriaOportunidad.Id, pvp, costo, costoIndirecto);

                        //Actualizar valores de la categoria
                        ActualizarCategoria(_service, bit_CategoriaOportunidad.Id, pvp, costo, costoIndirecto);

                        //Actualizar total oportunidad
                        ActualizarValoresOportunidad(_service, opportunidadCrm.Id, pvp, costo, costoIndirecto);

                    }
                }
                else
                {                    
                    SystemUser comercial = new SystemUser();
                    bit_modelo_comercial modelo = new bit_modelo_comercial();
                    bit_unidaddenegocio unidad = new bit_unidaddenegocio();

                    //Buscar comercial
                    comercial = ctx.SystemUserSet.Where(cm => cm.DomainName.Equals("admin-crm@grupobusiness.it")).FirstOrDefault();

                    //Buscar modelo comercial
                    modelo = ctx.bit_modelo_comercialSet.Where(mc => mc.bit_name.Equals("SERVICIO ADMINISTRADO")).FirstOrDefault();

                    //Buscar unidad negocio
                    unidad = ctx.bit_unidaddenegocioSet.Where(u => u.bit_name.Equals("GESTION DOCUMENTAL")).FirstOrDefault();

                    Opportunity op = new Opportunity
                    {
                        Name = "Servicios Cloud Azure - Icaro",
                        ParentAccountId = new EntityReference(Account.EntityLogicalName, cliente.Id),
                        OwnerId = new EntityReference(SystemUser.EntityLogicalName, comercial.Id),
                        EstimatedCloseDate = DateTime.Today,
                        OpportunityRatingCode = new OptionSetValue(1), // ALTA 1 MEDIA 2 BAJA 3
                        bit_cuantos_productos = 1,
                        bit_modelo_comercial = new EntityReference(bit_modelo_comercial.EntityLogicalName, modelo.Id),
                        new_UnidadNegocio = new EntityReference(bit_unidaddenegocio.EntityLogicalName, unidad.Id)
                    };
                    var response = _service.Create(op);
                    

                    //Crear categoria 
                    Guid idCategoriaOportunidad = CrearCategoria(_service, comercial.Id, response, bit_CatalogoCategoria.Id);

                    //Crear el item de la factura - crear por cada detalle de Icaro
                    CrearDetallePago(_service, comercial.Id, response, bit_CatalogoCategoria.Id, idCategoriaOportunidad, pvp, costo, costoIndirecto);

                    //Actualizar valores de la categoria
                    ActualizarCategoria(_service, idCategoriaOportunidad, pvp, costo, costoIndirecto);

                    //Actualizar total oportunidad
                    ActualizarValoresOportunidad(_service, response, pvp, costo, costoIndirecto);

                    //Cambiar al 60% la oportunidad
                    ChangeStage(_service, response);
                }


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

        public static Guid CrearDetallePago(IOrganizationService _service, Guid comercial, Guid oportunidad, Guid descCategoria, Guid idCategoria, Money pvp, Money costo, Money costoIndirecto)
        {
            string methodName = "[static Guid CrearDetallePago(IOrganizationService _service, Guid comercial, Guid oportunidad, Guid categoria)]";
            try
            {
                //Crear el item de la factura
                bit_detallepago bit_Detallepago = new bit_detallepago
                {
                    OwnerId = new EntityReference(SystemUser.EntityLogicalName, comercial),
                    bit_name = "Aqui debemos especificar el nombre del centro de costo de Icaro",
                    bit_Oportunidad = new EntityReference(Opportunity.EntityLogicalName, oportunidad),
                    bit_categoria = new EntityReference(bit_categoria.EntityLogicalName, descCategoria),
                    bit_pago_asociado = new EntityReference(bit_pago.EntityLogicalName, idCategoria),
                    bit_descripcion_facturacion = "Descripcion de la factura",
                    bit_Valoringreso = pvp, // PVP
                    bit_Costo = costo,
                    bit_costo_indirecto = costoIndirecto,
                    bit_Fechadefacturacion = DateTime.Today,
                    bit_Fechadepago = DateTime.Today
                };
                return _service.Create(bit_Detallepago);
            }
            catch (Exception ex)
            {
                throw new MsgException("Error al crear detalle de pago", m_className, methodName, ex);
            }
        }

        public static Guid CrearCategoria(IOrganizationService _service, Guid comercial, Guid oportunidad, Guid categoria)
        {
            string methodName = "[static Guid CrearDetallePago(IOrganizationService _service, Guid comercial, Guid oportunidad, Guid categoria)]";
            try
            {
                bit_pago bitPago = new bit_pago()
                {
                    OwnerId = new EntityReference(SystemUser.EntityLogicalName, comercial),
                    bit_Opotunidad = new EntityReference(Opportunity.EntityLogicalName, oportunidad),
                    bit_categoria_item = new EntityReference(bit_categoria.EntityLogicalName, categoria),
                    bit_descripcion_facturacion = "Descripcion de la factura",
                    bit_Numerodepagos = 0,
                    bit_PVP = new Money(0),
                    bit_Costo = new Money(0),
                    bit_costo_indirecto = new Money(0)
                };

                return _service.Create(bitPago);
            }
            catch (Exception ex)
            {
                throw new MsgException("Error al crear categoria", m_className, methodName, ex);
            }
        }

        public static void ActualizarCategoria(IOrganizationService _service, Guid idcategoria, Money pvp, Money costo, Money costoIndirecto)
        {
            string methodName = "[static Guid CrearDetallePago(IOrganizationService _service, Guid comercial, Guid oportunidad, Guid categoria)]";
            try
            {
                //Actualiza los valores del detalle de la categoria
                bit_pago bit_PagoUdp = new bit_pago
                {
                    Id = idcategoria,
                    bit_PVP = pvp,
                    bit_Costo = costo,
                    bit_costo_indirecto = costoIndirecto
                };
                _service.Update(bit_PagoUdp);
            }
            catch (Exception ex)
            {
                throw new MsgException("Error al actualizar categoria", m_className, methodName, ex);
            }
        }

        public static void ActualizarValoresOportunidad(IOrganizationService _service, Guid idOportunidad, Money pvp, Money costo, Money costoIndirecto)
        {
            string methodName = "[static Guid CrearDetallePago(IOrganizationService _service, Guid comercial, Guid oportunidad, Guid categoria)]";
            try
            {
                //Actualiza los valores del detalle de la categoria
                Opportunity op = new Opportunity
                {
                    Id = idOportunidad,
                    bit_PVP = pvp,
                    bit_CostoTotal = costo.Value,
                    bit_costo_indirecto = costoIndirecto.Value
                };
                _service.Update(op);
            }
            catch (Exception ex)
            {
                throw new MsgException("Error al actualizar categoria", m_className, methodName, ex);
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


            // Retrieve the stage ID of the next stage that you want to set as active
            activeStageID = (Guid)pathResp.ProcessStages.Entities[4].Attributes["processstageid"];

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
