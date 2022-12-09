using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using Xrm.Framework.CI.Common.Entities;

namespace Xrm.Framework.CI.Common.Common
{
    public class ComponentInfo
    {

        public string Name { get; set; }
        public string LogicalName { get; set; }
        public ComponentType ComponentType { get; set; }
        public Guid ObjectId { get; set; }
        public bool? IsManaged { get; set; }

        /// <summary>
        /// In which solutions this item is present
        /// </summary>
        public SolutionInfo[] Solutions { get; set; }

        ComponentInfo() {}

        public static ComponentInfo GetFromComponent(CIContext context, Guid objectId, ComponentType componentType)
        {
            var solutionManagementRepository = new SolutionManagementRepository(context);
            var result = new ComponentInfo()
            {
                ObjectId = objectId,
                ComponentType = componentType,
                Solutions = solutionManagementRepository.GetSolutionsContainingObject(objectId).Select(x => SolutionInfo.GetFromSolution(x)).ToArray()
            };
            try
            {
                var details = GetComponentDetails(context, componentType, objectId);
                result.IsManaged = details.isManaged;
                result.LogicalName = details.logicalName;
                result.Name = details.name;
            }
            catch (InvalidOperationException)
            {
                //InvalidOperationException thrown when failing to retrieve entity by id. Unfortunately, some components may exist as Dependency but not be a valid component. At least OnPrem instance had SdkMessageProcessingStep as a dependency, but object by itsellf doesn't exist.
            }
            catch (FaultException<OrganizationServiceFault> ex) when (ex.Detail.ErrorDetails.TryGetValue("ApiOriginalExceptionKey", out object originalException) && (originalException as string).StartsWith("Microsoft.Crm.BusinessEntities.CrmObjectNotFoundException"))
            {
                //Some metadata couldn't be retrieved
            }
            return result;
        }

        private static (bool? isManaged, string name, string logicalName) GetComponentDetails(CIContext context, ComponentType componentType, Guid objectId)
        {
            var solutionManagementRepository = new SolutionManagementRepository(context);
            switch (componentType)
            {
                case ComponentType.Attribute:
                    var attributeMetadata = context.GetAttributeMetadata(objectId);
                    return (attributeMetadata.IsManaged, $"{attributeMetadata.LogicalName} ({attributeMetadata.EntityLogicalName})", attributeMetadata.LogicalName);
                case ComponentType.Entity:
                    var entityMetadata1 = context.GetEntityMetadata(objectId);
                    return (entityMetadata1.IsManaged, entityMetadata1.LogicalName, entityMetadata1.LogicalName);
                case ComponentType.EntityRelationship:
                    var relationshipBase = context.GetRelationshipMetadata(objectId);
                    return (relationshipBase.IsManaged, $"{relationshipBase.SchemaName} ({relationshipBase.RelationshipType})", relationshipBase.SchemaName);
                case ComponentType.OptionSet:
                    var optionSetMetadata = context.GetOptionSetMetadata(objectId);
                    return (optionSetMetadata.IsManaged, optionSetMetadata.Name, optionSetMetadata.Name);
                case ComponentType.Workflow:
                    var workflow = solutionManagementRepository.GetEntityById<Workflow>(objectId, x => new Workflow() { WorkflowId = x.WorkflowId, Name = x.Name, ["ismanaged"] = x.IsManaged });
                    return (workflow.IsManaged, workflow.Name, workflow.LogicalName);
                case ComponentType.ConnectionRole:
                    var connectionRole = solutionManagementRepository.GetEntityById<ConnectionRole>(objectId, x => new ConnectionRole() { ConnectionRoleId = x.ConnectionRoleId, Name = x.Name, ["ismanaged"] = x.IsManaged });
                    return (connectionRole.IsManaged, connectionRole.Name, connectionRole.LogicalName);
                case ComponentType.SDKMessage:
                    var sdkMessage = solutionManagementRepository.GetEntityById<SdkMessage>(objectId, x => new SdkMessage() { SdkMessageId = x.SdkMessageId, Name = x.Name, ["ismanaged"] = x.IsManaged });
                    return (sdkMessage.IsManaged, sdkMessage.Name, sdkMessage.LogicalName);
                case ComponentType.SDKMessageProcessingStep:
                    var sdkMessageProcessingStep = solutionManagementRepository.GetEntityById<SdkMessageProcessingStep>(objectId, x => new SdkMessageProcessingStep() { SdkMessageProcessingStepId = x.SdkMessageProcessingStepId, Name = x.Name, ["ismanaged"] = x.IsManaged });
                    return (sdkMessageProcessingStep.IsManaged, sdkMessageProcessingStep.Name, sdkMessageProcessingStep.LogicalName);
                case ComponentType.SDKMessageProcessingStepImage:
                    var sdkMessageProcessingStepImage = solutionManagementRepository.GetEntityById<SdkMessageProcessingStepImage>(objectId, x => new SdkMessageProcessingStepImage() { SdkMessageProcessingStepImageId = x.SdkMessageProcessingStepImageId, Name = x.Name, ["ismanaged"] = x.IsManaged });
                    return (sdkMessageProcessingStepImage.IsManaged, sdkMessageProcessingStepImage.Name, sdkMessageProcessingStepImage.LogicalName);
                case ComponentType.PluginType:
                    var pluginType = solutionManagementRepository.GetEntityById<PluginType>(objectId, x => new PluginType() { PluginTypeId = x.PluginTypeId, Name = x.Name, ["ismanaged"] = x.IsManaged });
                    return (pluginType.IsManaged, pluginType.Name, pluginType.LogicalName);
                case ComponentType.PluginAssembly:
                    var pluginAssembly = solutionManagementRepository.GetEntityById<PluginAssembly>(objectId, x => new PluginAssembly() { PluginAssemblyId = x.PluginAssemblyId, Name = x.Name, ["ismanaged"] = x.IsManaged });
                    return (pluginAssembly.IsManaged, pluginAssembly.Name, pluginAssembly.LogicalName);
                case ComponentType.Role:
                    var role = solutionManagementRepository.GetEntityById<Role>(objectId, x => new Role() { RoleId = x.RoleId, Name = x.Name, ["ismanaged"] = x.IsManaged });
                    return (role.IsManaged, role.Name, role.LogicalName);
                case ComponentType.SavedQuery:
                    var savedQuery = solutionManagementRepository.GetEntityById<SavedQuery>(objectId, x => new SavedQuery() { SavedQueryId = x.SavedQueryId, Name = x.Name, ["ismanaged"] = x.IsManaged }) ;
                    return (savedQuery.IsManaged, savedQuery.Name, savedQuery.LogicalName);
                case ComponentType.ServiceEndpoint:
                    var serviceEndpoint = solutionManagementRepository.GetEntityById<ServiceEndpoint>(objectId, x => new ServiceEndpoint() { ServiceEndpointId = x.ServiceEndpointId, Name = x.Name, ["ismanaged"] = x.IsManaged });
                    return (serviceEndpoint.IsManaged, serviceEndpoint.Name, serviceEndpoint.LogicalName);
                case ComponentType.SystemForm:
                    var systemForm = solutionManagementRepository.GetEntityById<SystemForm>(objectId, x => new SystemForm() { FormId = x.FormId, Name = x.Name, ["type"] = x.Type, ["ismanaged"] = x.IsManaged });
                    return (systemForm.IsManaged, $"{systemForm.Name} ({systemForm.TypeEnum})", systemForm.LogicalName);
                case ComponentType.WebResource:
                    var webResource = solutionManagementRepository.GetEntityById<WebResource>(objectId, x => new WebResource() { WebResourceId = x.WebResourceId, Name = x.Name, ["ismanaged"] = x.IsManaged });
                    return (webResource.IsManaged, webResource.Name, webResource.LogicalName);
                case ComponentType.Report:
                    var report = solutionManagementRepository.GetEntityById<Report>(objectId, x => new Report() { ReportId = x.ReportId, Name = x.Name, ["ismanaged"] = x.IsManaged });
                    return (report.IsManaged, report.Name, report.LogicalName);
                case ComponentType.ContractTemplate:
                    var contractTemplate = solutionManagementRepository.GetEntityById<ContractTemplate>(objectId, x => new ContractTemplate() { ContractTemplateId = x.ContractTemplateId, Name = x.Name, ["ismanaged"] = x.IsManaged });
                    return (contractTemplate.IsManaged, contractTemplate.Name, contractTemplate.LogicalName);
                case ComponentType.EmailTemplate:
                    var template = solutionManagementRepository.GetEntityById<Template>(objectId, x => new Template() { TemplateId = x.TemplateId, Title = x.Title, ["ismanaged"] = x.IsManaged });
                    return (template.IsManaged, template.Title, template.LogicalName);
                case ComponentType.KBArticleTemplate:
                    var kbArticleTemplate = solutionManagementRepository.GetEntityById<KbArticleTemplate>(objectId, x => new KbArticleTemplate() { KbArticleTemplateId = x.KbArticleTemplateId, Title = x.Title, ["ismanaged"] = x.IsManaged });
                    return (kbArticleTemplate.IsManaged, kbArticleTemplate.Title, kbArticleTemplate.LogicalName);
                case ComponentType.RibbonCustomization:
                    var ribbonCustomization = solutionManagementRepository.GetEntityById<RibbonCustomization>(objectId, x => new RibbonCustomization() { RibbonCustomizationId = x.RibbonCustomizationId, Entity = x.Entity, ["ismanaged"] = x.IsManaged });
                    return (ribbonCustomization.IsManaged, $"Ribbon {ribbonCustomization.Entity}", ribbonCustomization.LogicalName);
                case ComponentType.SiteMap:
                    var siteMap = solutionManagementRepository.GetEntityById<SiteMap>(objectId, x => new SiteMap() { Id = x.Id, SiteMapNameUnique = x.SiteMapNameUnique, ["ismanaged"] = x.IsManaged });
                    return (siteMap.IsManaged, siteMap.SiteMapNameUnique, siteMap.LogicalName);
                case ComponentType.MailMergeTemplate:
                    var mailMergeTemplate = solutionManagementRepository.GetEntityById<MailMergeTemplate>(objectId, x => new MailMergeTemplate() { MailMergeTemplateId = x.MailMergeTemplateId, Name = x.Name, ["ismanaged"] = x.IsManaged });
                    return (mailMergeTemplate.IsManaged, mailMergeTemplate.Name, mailMergeTemplate.LogicalName);
                case ComponentType.SLA:
                    var sla = solutionManagementRepository.GetEntityById<SLA>(objectId, x => new SLA() { SLAId = x.SLAId, Name = x.Name, ["ismanaged"] = x.IsManaged });
                    return (sla.IsManaged, sla.Name, sla.LogicalName);
                case ComponentType.CustomControl:
                    var customControl = solutionManagementRepository.GetEntityById<CustomControl>(objectId, x => new CustomControl() { CustomControlId = x.CustomControlId, Name = x.Name, ["ismanaged"] = x.IsManaged });
                    return (customControl.IsManaged, customControl.Name, customControl.LogicalName);
                case ComponentType.FieldSecurityProfile:
                    var fieldSecurityProfile = solutionManagementRepository.GetEntityById<FieldSecurityProfile>(objectId, x => new FieldSecurityProfile() { FieldSecurityProfileId = x.FieldSecurityProfileId, Name = x.Name, ["ismanaged"] = x.IsManaged });
                    return (fieldSecurityProfile.IsManaged, fieldSecurityProfile.Name, fieldSecurityProfile.LogicalName);
                default:
                    return (null, $"{componentType} {objectId}", null);
            }
        }

        public override string ToString() => $"{ComponentType}: {Name}";
    }
}
