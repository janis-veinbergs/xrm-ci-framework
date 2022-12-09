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
        public SolutionComponent SolutionComponent { get; set; }

        /// <summary>
        /// Only will process unmanaged components. That is, managed ones will be skipped.
        /// </summary>
        [Parameter(ParameterSetName = RemoveParameterSetName, Mandatory = false)]
        [Parameter(ParameterSetName = DeleteParameterSetName, Mandatory = false)]
        public SwitchParameter Unmanaged { get; set; }
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
                ProcessDelete(SolutionComponent.ObjectId.Value, (ComponentType)SolutionComponent.ComponentType.Value);
            } else
            {
                Guid solutionId = GetSolutionId(context);

                if (SolutionComponent == null)
                {
                    base.WriteVerbose($"Removing all components from solution with Id : {solutionId}");
                    var querySolutionComponents = from s in context.SolutionComponentSet
                                                    where s.SolutionId == new EntityReference(Solution.EntityLogicalName, solutionId) && s.RootSolutionComponentId == null
                                                    select new { s.ComponentType, s.ObjectId };
                    querySolutionComponents.ToList().ForEach(x => ProcessRemove(x.ObjectId.Value, x.ComponentType));
                }
                else
                {
                    base.WriteVerbose($"Removing component {SolutionComponent.ObjectId} and Type: {(ComponentType)SolutionComponent.ComponentType.Value} from solution with Id : {solutionId}");
                    var queryOneSolutionComponent = from s in context.SolutionComponentSet
                                                    where s.SolutionId == new EntityReference(Solution.EntityLogicalName, solutionId) && s.ObjectId == SolutionComponent.ObjectId
                                                    select new { s.ComponentType, s.ObjectId };
                    var solutionComponent = queryOneSolutionComponent.FirstOrDefault();
                    if (solutionComponent != null)
                    {
                        ProcessRemove(solutionComponent.ObjectId.Value, solutionComponent.ComponentType);
                    }
                    else
                    {
                        WriteWarning($"Component {SolutionComponent.ObjectId} {(ComponentType)SolutionComponent.ComponentType.Value} not found within solution with Id : {solutionId}");
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
            DeleteObjectWithDependencies(objectId, componentType);
            base.WriteVerbose($"Deleted component with Id : {objectId} and Type: {componentType}");
        }

        public void DeleteObjectWithDependencies(Guid objectId, ComponentType? componentType, HashSet<string> deletingHashSet = null, int depth = 0)
        {
            if (deletingHashSet == null)
            {
                deletingHashSet = new HashSet<string>();
            }
            var objectkey = $"{componentType}{objectId}";
            if (deletingHashSet.Contains(objectkey))
            {
                return;
            }
            deletingHashSet.Add(objectkey);

            Logger.LogVerbose($"Checking dependencies for {componentType} / {objectId}");
            foreach (var objectToDelete in solutionManagementRepository.GetDependeciesForDelete(objectId, componentType.Value))
            {
                DeleteObjectWithDependencies(objectToDelete.DependentComponentObjectId.Value, objectToDelete.DependentComponentTypeEnum, deletingHashSet, ++depth);
            }

            switch (componentType)
            {
                case ComponentType.Entity:
                    var entityMetadata1 = OrganizationService.GetEntityMetadata(objectId);
                    if (!Unmanaged.IsPresent || (Unmanaged.IsPresent && !entityMetadata1.IsManaged.Value))
                    {
                        if (ShouldProcess($"{componentType} {entityMetadata1.LogicalName} {objectId}"))
                        {
                            Logger.LogVerbose($"Trying to delete {componentType} {entityMetadata1.LogicalName} {objectId}");
                            OrganizationService.Execute(new DeleteEntityRequest { LogicalName = entityMetadata1.LogicalName });
                        }
                    }
                    else
                    {
                        Logger.LogInformation($"Skipping Managed {componentType} {entityMetadata1.LogicalName} {objectId}");
                    }
                    break;
                case ComponentType.EntityRelationship:
                    Logger.LogVerbose($"Trying to delete {nameof(ComponentType.EntityRelationship)} {componentType} {objectId}");
                    var relationshipBase = OrganizationService.GetRelationshipMetadata(objectId);
                    if (!Unmanaged.IsPresent || (Unmanaged.IsPresent && !relationshipBase.IsManaged.Value))
                    {
                        Logger.LogVerbose($"Checking dependencies for {relationshipBase.RelationshipType} {relationshipBase.SchemaName} {objectId}");
                        DeleteObjectWithDependencies(objectId, ComponentType.EntityRelationship, deletingHashSet, ++depth);
                        if (ShouldProcess($"{relationshipBase.RelationshipType} {relationshipBase.SchemaName}"))
                            OrganizationService.Execute(new DeleteRelationshipRequest()
                            {
                                Name = relationshipBase.SchemaName
                            });
                    }
                    else
                    {
                        Logger.LogInformation($"Skipping Managed {componentType} {relationshipBase.RelationshipType}: {relationshipBase.SchemaName} {objectId}");
                    }
                    break;
                case ComponentType.OptionSet:
                    var optionSetMetadata = solutionManagementRepository.GetOptionSetMetadata(objectId);
                    if (!Unmanaged.IsPresent || (Unmanaged.IsPresent && !optionSetMetadata.IsManaged.Value))
                    {
                        Logger.LogVerbose($"Deleting {nameof(componentType)} {optionSetMetadata.Name} {objectId}");
                        if (ShouldProcess($"OptionSet {optionSetMetadata.Name}"))
                        {
                            OrganizationService.Execute(new DeleteOptionSetRequest()
                            {
                                Name = optionSetMetadata.Name
                            });
                        }
                    }
                    else
                    {
                        Logger.LogInformation($"Skipping Managed {nameof(componentType)} {optionSetMetadata.Name} {objectId}");
                    }
                    break;
                case ComponentType.Workflow:
                    var workflow = solutionManagementRepository.GetEntityById<Workflow>(objectId, x => new Workflow()
                    {
                        Name = x.Name,
                        UniqueName = x.UniqueName,
                        StateCode = x.StateCode,
                        Category = x.Category,
                        WorkflowId = x.WorkflowId,
                        ["ismanaged"] = x.IsManaged,
                    });
                    if (!Unmanaged.IsPresent || (Unmanaged.IsPresent && !workflow.IsManaged.Value))
                    {
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
                            DeleteObjectWithDependencies(entityMetadata2.MetadataId.Value, ComponentType.Entity, deletingHashSet, ++depth);
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
                        if (ShouldProcess(GetEntityComponentName(workflow)))
                        {
                            OrganizationService.Delete(Workflow.EntityLogicalName, objectId);
                        }
                    }
                    else
                    {
                        Logger.LogInformation($"Skipping Managed {nameof(componentType)} {workflow.Name} {objectId}");
                    }
                    break;
                case ComponentType.ConnectionRole:
                    DeleteEntityObject<ConnectionRole>(objectId);
                    break;
                case ComponentType.SDKMessageProcessingStep:
                    DeleteEntityObject<SdkMessageProcessingStep>(objectId);
                    break;
                case ComponentType.SDKMessageProcessingStepImage:
                    DeleteEntityObject<SdkMessageProcessingStepImage>(objectId);
                    break;
                case ComponentType.PluginType:
                    DeleteEntityObject<PluginType>(objectId);
                    break;
                case ComponentType.PluginAssembly:
                    DeleteEntityObject<PluginAssembly>(objectId);
                    break;
                case ComponentType.ServiceEndpoint:
                    DeleteEntityObject<ServiceEndpoint>(objectId);
                    break;
                case ComponentType.SavedQuery:
                    DeleteEntityObject<SavedQuery>(objectId);
                    break;
                case ComponentType.WebResource:
                    DeleteEntityObject<WebResource>(objectId);
                    break;
                default:
                    Logger.LogWarning($"Cannot delete {componentType} {objectId}: Delete for component type not implemented.");
                    break;
            }
        }

        private void DeleteEntityObject<TEntity>(Guid entityId) where TEntity: Entity
        {
            Entity entity;
            try
            {
                entity = solutionManagementRepository.GetEntityById<TEntity>(entityId);
            } catch (InvalidOperationException)
            {
                //Entity not found
                Logger.LogWarning($"Entity {typeof(TEntity).Name} {entityId} not found, skipping");
                return;
            }
            var componentName = GetEntityComponentName(entity);
            if (ShouldProcess(componentName))
            {
                Logger.LogVerbose($"Deleting {componentName}");
                OrganizationService.Delete(entity.LogicalName, entity.Id);
            } else
            {
                Logger.LogInformation($"Skipping {componentName}");
            }
        }

        bool ShouldProcessEntityComponent<T>(T entity) where T : Entity => Unmanaged.IsPresent && entity.TryGetAttributeValue("ismanaged", out bool isManaged) && !isManaged && ShouldProcess(GetEntityComponentName(entity));

        string GetEntityComponentName<T>(T entity) where T : Entity => $"{entity.LogicalName} \"{GetEntityPrimaryNameFieldValue(entity)}\" {entity.Id}";

        string GetEntityPrimaryNameFieldValue<T>(T entity) where T : Entity
        {
            string primaryNameAttribute;
            try
            {
                primaryNameAttribute = (string)entity.GetType().GetField("PrimaryNameAttribute").GetValue(entity);
                if (primaryNameAttribute == null)
                {
                    return "";
                }
            } catch 
            {
                return "";
            }
            return entity.GetAttributeValue<string>(primaryNameAttribute);
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