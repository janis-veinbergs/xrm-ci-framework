using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using Xrm.Framework.CI.Common;
using Xrm.Framework.CI.Common.Entities;
using System.Reflection;
using Xrm.Framework.CI.Common.Common;

namespace Xrm.Framework.CI.PowerShell.Cmdlets
{
    /// <summary>
    /// <para type="synopsis">Removes a CRM Solution Components.</para>
    /// <para type="description">The Remove-XrmSolutionComponents of a CRM solution by unique name. Removes components from solution and does not delete them from system.
    /// </para>
    /// </summary>
    /// <example>
    ///   <code>C:\PS>Remove-XrmSolutionComponents -ConnectionString "" -UniqueSolutionName "UniqueSolutionName"</code>
    ///   <para>Exports the "" managed solution to "" location</para>
    /// </example>
    [Cmdlet(VerbsCommon.Remove, "XrmSolutionComponents", SupportsShouldProcess = true, DefaultParameterSetName = RemoveParameterSetName)]
    [OutputType(typeof(String))]
    public class RemoveXrmSolutionComponentsCommand : XrmCommandBase
    {
        private const string RemoveParameterSetName = "Remove";
        private const string DeleteParameterSetName = "Delete";
        #region Parameters

        /// <summary>
        /// <para type="description">The unique name of the solution components to be removed.</para>
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = RemoveParameterSetName)]
        public string SolutionName { get; set; }

        /// <summary>
        /// Delete component from system. It is a destructive action!
        /// </summary>
        [Parameter(ParameterSetName = DeleteParameterSetName)]
        public SwitchParameter Delete { get; set; }


        /// <summary>
        /// Removes only single solution component.
        /// </summary>
        [Parameter(ParameterSetName = RemoveParameterSetName, ValueFromPipeline = true, Mandatory = false)]
        [Parameter(ParameterSetName = DeleteParameterSetName, ValueFromPipeline = true, Mandatory = true)]
        public ComponentInfo ComponentInfo { get; set; }
        #endregion

        CIContext context;
        SolutionComponentsManager solutionComponentsManager;
        SolutionManagementRepository solutionManagementRepository;
        #region Process Record

        protected override void BeginProcessing()
        {
            base.BeginProcessing();
            context = new CIContext(OrganizationService);
            solutionComponentsManager = new SolutionComponentsManager(Logger, OrganizationService);
            solutionManagementRepository = new SolutionManagementRepository(context);
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            base.WriteVerbose(string.Format("Removing Solution Components: {0}", SolutionName));

            if (Delete.IsPresent)
            {
                ProcessDelete(ComponentInfo.ObjectId, ComponentInfo.ComponentType);
            } else
            {
                Guid solutionId = GetSolutionId(context);

                if (ComponentInfo == null)
                {
                    base.WriteVerbose($"Removing all components from solution with Id : {solutionId}");
                    var querySolutionComponents = from s in context.SolutionComponentSet
                                                    where s.SolutionId == new EntityReference(Solution.EntityLogicalName, solutionId) && s.RootSolutionComponentId == null
                                                    select new { s.ComponentType, s.ObjectId };
                    querySolutionComponents.ToList().ForEach(x => ProcessRemove(x.ObjectId.Value, x.ComponentType));
                }
                else
                {
                    base.WriteVerbose($"Removing component {ComponentInfo.ObjectId} and Type: {ComponentInfo.ComponentType} from solution with Id : {solutionId}");
                    var queryOneSolutionComponent = from s in context.SolutionComponentSet
                                                    where s.SolutionId == new EntityReference(Solution.EntityLogicalName, solutionId) && s.ObjectId == ComponentInfo.ObjectId
                                                    select new { s.ComponentType, s.ObjectId };
                    var solutionComponent = queryOneSolutionComponent.FirstOrDefault();
                    if (solutionComponent != null)
                    {
                        ProcessRemove(solutionComponent.ObjectId.Value, solutionComponent.ComponentType);
                    }
                    else
                    {
                        WriteWarning($"Component {ComponentInfo.ObjectId} {ComponentInfo.ComponentType} not found within solution with Id : {solutionId}");
                    }
                }
            }
        }

        protected override void EndProcessing()
        {
            if (context != null)
            {
                context.Dispose();
                context = null;
            }
        }

        private Guid GetSolutionId(CIContext context)
        {
            //Process removal of ALL components within Solution
            var querySolution = from solution in context.SolutionSet
                                where solution.UniqueName == SolutionName
                                select solution.Id;
            var solutionId = querySolution.FirstOrDefault();
            if (solutionId == default(Guid))
            {
                throw new Exception(string.Format("Solution {0} could not be found", SolutionName));
            }

            return solutionId;
        }

        void ProcessDelete(Guid objectId, ComponentType componentType)
        {
            if (!Delete.IsPresent) { throw new InvalidOperationException($"{nameof(ProcessDelete)} not callable without {nameof(Delete)} switch parameter set"); }
            DeleteObject(ComponentInfo.GetFromComponent(context, objectId, componentType));
            base.WriteVerbose($"Deleted component with Id : {objectId} and Type: {componentType}");
        }

        public void DeleteObject(ComponentInfo component)
        {
            Guid objectId = component.ObjectId;
            ComponentType componentType = component.ComponentType;

            switch (componentType)
            {
                case ComponentType.Entity:
                    var entityMetadata1 = OrganizationService.GetEntityMetadata(objectId);
                    if (ShouldProcess($"{componentType} {entityMetadata1.LogicalName} {objectId}"))
                    {
                        Logger.LogVerbose($"Trying to delete {componentType} {entityMetadata1.LogicalName} {objectId}");
                        OrganizationService.Execute(new DeleteEntityRequest { LogicalName = entityMetadata1.LogicalName });
                    }
                    break;
                case ComponentType.EntityRelationship:
                    Logger.LogVerbose($"Trying to delete {nameof(ComponentType.EntityRelationship)} {componentType} {objectId}");
                    var relationshipBase = OrganizationService.GetRelationshipMetadata(objectId);
                    if (ShouldProcess($"{relationshipBase.RelationshipType} {relationshipBase.SchemaName}"))
                        OrganizationService.Execute(new DeleteRelationshipRequest()
                        {
                            Name = relationshipBase.SchemaName
                        });
                    break;
                case ComponentType.OptionSet:
                    var optionSetMetadata = solutionManagementRepository.GetOptionSetMetadata(objectId);
                    Logger.LogVerbose($"Deleting {nameof(componentType)} {optionSetMetadata.Name} {objectId}");
                    if (ShouldProcess($"OptionSet {optionSetMetadata.Name}"))
                    {
                        OrganizationService.Execute(new DeleteOptionSetRequest()
                        {
                            Name = optionSetMetadata.Name
                        });
                    }
                    break;
                case ComponentType.Workflow:
                    Workflow workflow = null;
                    try
                    {
                        workflow = solutionManagementRepository.GetEntityById<Workflow>(objectId, x => new Workflow()
                        {
                            Name = x.Name,
                            UniqueName = x.UniqueName,
                            StateCode = x.StateCode,
                            Category = x.Category,
                            WorkflowId = x.WorkflowId,
                            ["ismanaged"] = x.IsManaged,
                        });
                    } catch (InvalidOperationException ex)
                    {
                        WriteError(new ErrorRecord(ex, ex.Source, ErrorCategory.InvalidOperation, component));
                    }
                    if (workflow.StateCode == WorkflowState.Activated)
                    {
                        Logger.LogVerbose($"Unpublishing workflow {workflow.Name}");
                        if (ShouldProcess($"Deactivate {workflow.Name}"))
                        {
                            OrganizationService.Execute(new SetStateRequest
                            {
                                EntityMoniker = workflow.ToEntityReference(),
                                State = new OptionSetValue((int)WorkflowState.Draft),
                                Status = new OptionSetValue((int)Workflow_StatusCode.Draft)
                            });
                        }
                    }
                    if (workflow.CategoryEnum == Workflow_Category.BusinessProcessFlow)
                    {
                        var entityMetadata2 = OrganizationService.GetEntityMetadata(workflow.UniqueName);
                        Logger.LogVerbose($"Checking dependencies for BPF entity: {workflow.UniqueName}");
                        DeleteObject(ComponentInfo.GetFromComponent(context, entityMetadata2.MetadataId.Value, ComponentType.Entity));
                    }

                    if (workflow.CategoryEnum == Workflow_Category.BusinessProcessFlow)
                    {
                        if (ShouldProcess($"Business Process Flow {workflow.Name}"))
                        {
                            RemoveAllWorkflowsFromBpf(workflow);
                        }
                        Logger.LogVerbose($"Preserving BPF {workflow.Name}");
                        return;
                    }

                    Logger.LogVerbose($"Trying to delete {componentType} {workflow.Name}");
                    if (ShouldProcess(ComponentInfo.Name))
                    {
                        OrganizationService.Delete(Workflow.EntityLogicalName, objectId);
                    }
                    break;
                case ComponentType.ConnectionRole:
                case ComponentType.SDKMessage:
                case ComponentType.SDKMessageProcessingStep:
                case ComponentType.SDKMessageProcessingStepImage:
                case ComponentType.PluginType:
                case ComponentType.PluginAssembly:
                case ComponentType.Role:
                case ComponentType.SavedQuery:
                case ComponentType.ServiceEndpoint:
                case ComponentType.SystemForm:
                case ComponentType.WebResource:
                case ComponentType.Report:
                case ComponentType.ContractTemplate:
                case ComponentType.EmailTemplate:
                case ComponentType.KBArticleTemplate:
                case ComponentType.RibbonCustomization:
                case ComponentType.SiteMap:
                case ComponentType.MailMergeTemplate:
                case ComponentType.SLA:
                case ComponentType.CustomControl:
                case ComponentType.FieldSecurityProfile:
                    DeleteEntityObject(component);
                    break;
                default:
                    Logger.LogWarning($"Cannot delete {componentType} {objectId}: Delete for component type not implemented.");
                    break;
            }
        }

        private void DeleteEntityObject(ComponentInfo component)
        {
            if (component.LogicalName == null)
            {
                throw new ArgumentException($"Component {component} has empty {nameof(component.LogicalName)}"); ;
            }
            if (ShouldProcess(component.Name))
            {
                Logger.LogVerbose($"Deleting {component}");
                OrganizationService.Delete(component.LogicalName, component.ObjectId);
            }
            else
            {
                Logger.LogInformation($"Skipping {component}");
            }
        }

        void ProcessRemove(Guid componentId, OptionSetValue componentType)
        {
            if (Delete.IsPresent) { throw new InvalidOperationException($"{nameof(ProcessRemove)} not callable with {nameof(Delete)} switch parameter set"); }
            var componentTypeEnum = (ComponentType)componentType.Value;
            if (ShouldProcess($"{componentTypeEnum} {componentId} from Solution {SolutionName}"))
            {
                var removeReq = new RemoveSolutionComponentRequest()
                {
                    ComponentId = componentId,
                    ComponentType = componentType.Value,
                    SolutionUniqueName = SolutionName
                };
                OrganizationService.Execute(removeReq);
            }
            base.WriteVerbose($"Removed component from solution {SolutionName} with Id : {componentId} and Type: {componentTypeEnum}");
        }

        private const string ActionComposieClassWithAssemblyQualifiedName = "Microsoft.Crm.Workflow.Activities.ActionComposite, Microsoft.Crm.Workflow, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35";
        private const string mxswaNamespace = "clr-namespace:Microsoft.Xrm.Sdk.Workflow.Activities;assembly=Microsoft.Xrm.Sdk.Workflow, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35";

        private void RemoveAllWorkflowsFromBpf(Workflow bpf)
        {
            var xaml = XDocument.Parse(bpf.Xaml);
            var nsmgr = new XmlNamespaceManager(new NameTable());
            nsmgr.AddNamespace("mxswa", mxswaNamespace);
            var actionsElements = xaml.XPathSelectElements($"//mxswa:ActivityReference[@AssemblyQualifiedName='{ActionComposieClassWithAssemblyQualifiedName}']", nsmgr).ToList();
            foreach (var element in actionsElements)
            {
                element.Remove();
            }
            OrganizationService.Update(new Workflow
            {
                Xaml = xaml.ToString(SaveOptions.DisableFormatting),
                Id = bpf.Id
            });
        }
        #endregion
    } 
}