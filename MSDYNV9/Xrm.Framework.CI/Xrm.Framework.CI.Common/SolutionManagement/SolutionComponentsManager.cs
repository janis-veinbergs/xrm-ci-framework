using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using Xrm.Framework.CI.Common.Entities;
using Xrm.Framework.CI.Common.Logging;

namespace Xrm.Framework.CI.Common
{
    public class SolutionComponentsManager : XrmBase
    {
        #region Variables

        private const string ImportSuccess = "success";
        private const string ImportFailure = "failure";
        private const string ImportProcessed = "processed";
        private const string ImportUnprocessed = "Unprocessed";

        #endregion

        #region Properties

        protected IOrganizationService PollingOrganizationService
        {
            get;
            set;
        }
        
        readonly CIContext context;
        #endregion

        #region Constructors

        public SolutionComponentsManager(ILogger logger,
            IOrganizationService organizationService)
            : base(logger, organizationService)
        {
            context = new CIContext(OrganizationService);
        }

        #endregion

        #region Methods

        public MissingComponent[] GetMissingComponentsOnTarget(
                string solutionFilePath)
        {
            Logger.LogInformation("Retrieving Missing Components for  Solution: {0}", solutionFilePath);

            if (!File.Exists(solutionFilePath))
            {
                Logger.LogError("Solution File does not exist: {0}", solutionFilePath);
                throw new FileNotFoundException("Solution File does not exist", solutionFilePath);
            }

            SolutionXml solutionXml = new SolutionXml(Logger);

            XrmSolutionInfo info = solutionXml.GetSolutionInfoFromZip(solutionFilePath);

            if (info == null)
            {
                throw new Exception("Invalid Solution File");
            }
            else
            {
                Logger.LogInformation("Solution Unique Name: {0}, Version: {1}",
                    info.UniqueName,
                    info.Version);
            }

            byte[] solutionBytes = File.ReadAllBytes(solutionFilePath);

            var request = new RetrieveMissingComponentsRequest()
            {
                CustomizationFile = solutionBytes
            };

            RetrieveMissingComponentsResponse response = OrganizationService.Execute(request) as RetrieveMissingComponentsResponse;

            Logger.LogInformation("{0} Missing Components retrieved for Solution", response.MissingComponents.Length);

            return response.MissingComponents;
        }

        public EntityCollection GetMissingDependencies(
            string SolutionName)
        {
            Logger.LogInformation("Retrieving Missing Dependencies for Solution: {0}", SolutionName);

            if (string.IsNullOrEmpty(SolutionName))
            {
                throw new Exception("SolutionName is required to retrieve missing dependencies");
            }

            var request = new RetrieveMissingDependenciesRequest()
            {
                SolutionUniqueName = SolutionName
            };

            RetrieveMissingDependenciesResponse response = OrganizationService.Execute(request) as RetrieveMissingDependenciesResponse;

            Logger.LogInformation("{0} Missing dependencies retrieved for Solution", response.EntityCollection.Entities.Count);

            return response.EntityCollection;
        }

        PluginRepository pluginRepository = null;
        public void DeleteObjectWithDependencies(Guid objectId, ComponentType? componentType, HashSet<string> deletingHashSet = null)
        {
            if (deletingHashSet == null)
            {
                deletingHashSet = new HashSet<string>();
            }
            if (pluginRepository == null)
            {
                pluginRepository = new PluginRepository(context);
            }
            var objectkey = $"{componentType}{objectId}";
            if (deletingHashSet.Contains(objectkey))
            {
                return;
            }
            deletingHashSet.Add(objectkey);

            Logger.LogVerbose($"Checking dependencies for {componentType} / {objectId}");
            foreach (var objectToDelete in GetDependeciesForDelete(objectId, componentType))
            {
                DeleteObjectWithDependencies(objectToDelete.DependentComponentObjectId.Value, objectToDelete.DependentComponentTypeEnum, deletingHashSet);
            }

            switch (componentType)
            {
                case ComponentType.Entity:
                    var entityMetadata1 = OrganizationService.GetEntityMetadata(objectId);
                    Logger.LogVerbose($"Trying to delete {componentType} {entityMetadata1.LogicalName} {objectId}");
                    var deleterequest = new DeleteEntityRequest { LogicalName = entityMetadata1.LogicalName };
                    OrganizationService.Execute(deleterequest);
                    break;
                case ComponentType.Workflow:
                    var workflow = GetWorkflowById(objectId);
                    if (workflow.StateCode == WorkflowState.Activated)
                    {
                        Logger.LogVerbose($"Unpublishing workflow {workflow.Name}");
                        OrganizationService.Execute(new SetStateRequest
                        {
                            EntityMoniker = workflow.ToEntityReference(),
                            State = new OptionSetValue((int)WorkflowState.Draft),
                            Status = new OptionSetValue((int)Workflow_StatusCode.Draft)
                        });
                    }
                    if (workflow.CategoryEnum == Workflow_Category.BusinessProcessFlow)
                    {
                        var entityMetadata2 = OrganizationService.GetEntityMetadata(workflow.UniqueName);
                        Logger.LogVerbose($"Checking dependencies for BPF entity: {workflow.UniqueName}");
                        DeleteObjectWithDependencies(entityMetadata2.MetadataId.Value, ComponentType.Entity, deletingHashSet);
                    }

                    if (workflow.CategoryEnum == Workflow_Category.BusinessProcessFlow)
                    {
                        RemoveAllWorkflowsFromBpf(workflow);
                        Logger.LogVerbose($"Preserving BPF {workflow.Name}");
                        return;
                    }

                    Logger.LogVerbose($"Trying to delete {componentType} {workflow.Name}");
                    OrganizationService.Delete(Workflow.EntityLogicalName, objectId);
                    break;
                case ComponentType.SDKMessageProcessingStep:
                    var step = pluginRepository.GetSdkMessageProcessingStepById(objectId);
                    if (step?.IsHidden.Value == true)
                    {
                        Logger.LogVerbose($"Preserving hidden SdkMessageProcessingStep {step.Name}");
                        return;
                    }
                    Logger.LogVerbose($"Trying to delete {componentType} {step.Name} / {objectId}");
                    OrganizationService.Delete(SdkMessageProcessingStep.EntityLogicalName, objectId);
                    break;
                case ComponentType.SDKMessageProcessingStepImage:
                    Logger.LogVerbose($"Trying to delete {componentType} / {objectId}");
                    OrganizationService.Delete(SdkMessageProcessingStepImage.EntityLogicalName, objectId);
                    break;
                case ComponentType.PluginType:
                    var type = pluginRepository.GetPluginTypeById(objectId);
                    Logger.LogVerbose($"Trying to delete {componentType} {type.Name} / {objectId}");
                    OrganizationService.Delete(PluginType.EntityLogicalName, objectId);
                    break;
                case ComponentType.PluginAssembly:
                    Logger.LogVerbose($"Trying to delete {componentType} {objectId}");
                    OrganizationService.Delete(PluginAssembly.EntityLogicalName, objectId);
                    break;
                case ComponentType.ServiceEndpoint:
                    Logger.LogVerbose($"Trying to delete {componentType} {objectId}");
                    OrganizationService.Delete(ServiceEndpoint.EntityLogicalName, objectId);
                    break;
                case ComponentType.EntityRelationship:
                    Logger.LogVerbose($"Trying to delete {nameof(ComponentType.EntityRelationship)} {componentType} {objectId}");
                    var relationshipBase = OrganizationService.GetRelationshipMetadata(objectId);
                    Logger.LogVerbose($"Checking dependencies for {relationshipBase.RelationshipType}: {relationshipBase.SchemaName}");
                    DeleteObjectWithDependencies(objectId, ComponentType.EntityRelationship, deletingHashSet);
                    OrganizationService.Execute(new DeleteRelationshipRequest()
                    {
                        Name = relationshipBase.SchemaName
                    });
                    //switch (relationshipBase.RelationshipType)
                    //{
                    //    case Microsoft.Xrm.Sdk.Metadata.RelationshipType.OneToManyRelationship:
                    //        var oneToNRelationship = (OneToManyRelationshipMetadata)relationshipBase;

                    //        break;
                    //    case Microsoft.Xrm.Sdk.Metadata.RelationshipType.ManyToManyRelationship:
                    //        var NToNRelationship = (ManyToManyRelationshipMetadata)relationshipBase;
                    //        OrganizationService.Execute(new DeleteRelationshipRequest()
                    //        {
                    //            Name = oneToNRelationship.SchemaName
                    //        });
                    //        break;
                    //}
                    break;
                default:
                    Logger.LogWarning($"Cannot delete {componentType} {objectId}: Delete for component type not implemented.");
                    break;
            }
        }

        Workflow GetWorkflowById(Guid id) => new PluginRepository(context).GetWorkflowById(id);

        private IEnumerable<Dependency> GetDependeciesForDelete(Guid objectId, ComponentType? componentType) => ((RetrieveDependenciesForDeleteResponse)OrganizationService.Execute(new RetrieveDependenciesForDeleteRequest()
        {
            ComponentType = (int)componentType,
            ObjectId = objectId
        })).EntityCollection.Entities.Select(x => x.ToEntity<Dependency>());

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

        public void AddComponentToSolution(Guid componentId, ComponentType componentType, string solutionName)
        {
            if (string.IsNullOrEmpty(solutionName))
            {
                return;
            }

            Logger.LogVerbose($"Adding {componentType} {componentId} to solution {solutionName}");
            OrganizationService.Execute(new AddSolutionComponentRequest
            {
                AddRequiredComponents = false,
                ComponentId = componentId,
                ComponentType = (int)componentType,
                SolutionUniqueName = solutionName
            });
        }
        #endregion
    }


    //public Component
}
