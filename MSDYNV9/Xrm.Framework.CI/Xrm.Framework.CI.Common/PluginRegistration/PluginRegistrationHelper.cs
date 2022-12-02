using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Xrm.Framework.CI.Common.Entities;
using System.Xml.Linq;
using System.Xml;
using System.Xml.XPath;
using Xrm.Framework.CI.Common.Logging;
using Xrm.Framework.CI.Common;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace Xrm.Framework.CI.Common
{
    public class PluginRegistrationHelper
    {
        private readonly IOrganizationService organizationService;
        private readonly PluginRepository pluginRepository;
        private readonly Action<string> logVerbose;
        private readonly Action<string> logWarning;
        private readonly IReflectionLoader reflectionLoader;
        private readonly IPluginRegistrationObjectFactory pluginRegistrationObjectFactory;
        private readonly ILogger Logger;
        private readonly SolutionComponentsManager SolutionComponentsManager;

        public PluginRegistrationHelper(IOrganizationService service, CIContext xrmContext, Action<string> logVerbose, Action<string> logWarning)
        {
            this.logVerbose = logVerbose;
            this.logWarning = logWarning;
            this.Logger = new DelegateLogger(logError: null, logWarning: logWarning, logInformation: null, logVerbose: logVerbose);
            organizationService = service;
            pluginRepository = new PluginRepository(xrmContext);
            reflectionLoader = new ReflectionLoader();
            pluginRegistrationObjectFactory = new PluginRegistrationObjectFactory();
            SolutionComponentsManager = new SolutionComponentsManager(Logger, this.organizationService);
        }

        public PluginRegistrationHelper(Action<string> logVerbose, Action<string> logWarning)
        {
            this.logVerbose = logVerbose;
            this.logWarning = logWarning;
            this.Logger = new DelegateLogger(logError: null, logWarning: logWarning, logInformation: null, logVerbose: logVerbose);
            reflectionLoader = new ReflectionLoader();
            pluginRegistrationObjectFactory = new PluginRegistrationObjectFactory();
        }

        public PluginRegistrationHelper(Action<string> logVerbose, Action<string> logWarning,
            IReflectionLoader reflectionLoader, IPluginRegistrationObjectFactory pluginRegistrationObjectFactory)
        {
            this.logVerbose = logVerbose;
            this.logWarning = logWarning;
            this.Logger = new DelegateLogger(logError: null, logWarning: logWarning, logInformation: null, logVerbose: logVerbose);
            this.reflectionLoader = reflectionLoader;
            this.pluginRegistrationObjectFactory = pluginRegistrationObjectFactory;
            SolutionComponentsManager = new SolutionComponentsManager(Logger, this.organizationService);
        }

        public Assembly GetAssemblyRegistration(string assemblyName, string version) => pluginRepository.GetAssemblyRegistration(assemblyName, version);

        public void RemoveComponentsNotInMapping(Assembly assemblyMapping)
        {
            var assemblyInCrm = pluginRepository.GetAssemblyRegistration(assemblyMapping.Name, assemblyMapping.Version);
            if (assemblyInCrm == null)
            {
                logVerbose?.Invoke($"Assembly {assemblyMapping.Name} not found in CRM");
                return;
            }

            var stepsInMapping = new HashSet<string>(assemblyMapping.PluginTypes.SelectMany(t => t.Steps, (t, s) => $"{t.Name}#{s.Name}#{s.GetHashCode()}"));
            var pluginStepsToDelete = assemblyInCrm.PluginTypes.SelectMany(t => t.Steps, (t, s) => new
            {
                Key = $"{t.Name}#{s.Name}#{s.GetHashCode()}",
                Name = s.Name,
                Id = s.Id
            })
                .Where(x => !stepsInMapping.Contains(x.Key)).
                ToList();
            foreach (var pluginStep in pluginStepsToDelete)
            {
                logVerbose?.Invoke($"Trying to delete step {pluginStep.Id} / {pluginStep.Name}");
                organizationService.Delete(SdkMessageProcessingStep.EntityLogicalName, pluginStep.Id.Value);
            }

            var typesInMapping = new HashSet<string>(assemblyMapping.PluginTypes.Select(t => t.Name));
            var pluginTypesToDelete = assemblyInCrm.PluginTypes
                .Where(t => !typesInMapping.Contains(t.Name))
                .ToList();      
            foreach (var pluginType in pluginTypesToDelete)
            {
                logVerbose?.Invoke($"Trying to delete type {pluginType.Id} / {pluginType.Name}");
                SolutionComponentsManager.DeleteObjectWithDependencies(pluginType.Id.Value, ComponentType.PluginType);
            }
        }

        //Leaving this method within PluginRegistrationHelper, because the interface was public, thus not to break any dependant assemblies.
        public void DeleteObjectWithDependencies(Guid objectId, ComponentType? componentType, HashSet<string> deletingHashSet = null) => SolutionComponentsManager.DeleteObjectWithDependencies(objectId, componentType, deletingHashSet);
        public Guid UpsertPluginAssembly(Assembly pluginAssembly, AssemblyInfo assemblyInfo, string solutionName, RegistrationTypeEnum registrationType)
        {
            Guid Id = pluginAssembly?.Id ?? Guid.Empty;
            if (Id == Guid.Empty)
            {
                Id = pluginRepository.GetPluginAssemblyId(assemblyInfo.AssemblyName);
                logWarning?.Invoke($"Extracted id using plugin assembly name {assemblyInfo.AssemblyName}");
            }

            var assembly = new PluginAssembly()
            {
                Version = assemblyInfo.Version,
                Content = assemblyInfo.Content,
                Name = assemblyInfo.AssemblyName,
                SourceTypeEnum = PluginAssembly_SourceType.Database,
                IsolationModeEnum = PluginAssembly_IsolationMode.Sandbox,
            };

            if (pluginAssembly != null)
            {
                assembly.SourceTypeEnum = pluginAssembly.SourceType;
                assembly.IsolationModeEnum = pluginAssembly.IsolationMode;
            }

            if (!Id.Equals(Guid.Empty) && registrationType == RegistrationTypeEnum.Reset)
            {
                SolutionComponentsManager.DeleteObjectWithDependencies(Id, ComponentType.PluginAssembly);
            }

            logVerbose?.Invoke($"Trying to upsert {assemblyInfo.AssemblyName} / {Id}");
            Id = ExecuteRequest(registrationType, Id, assembly);

            SolutionComponentsManager.AddComponentToSolution(Id, ComponentType.PluginAssembly, solutionName);

            return Id;
        }

        public void UpsertPluginTypeAndSteps(Guid parentId, Type pluginType, string solutionName, RegistrationTypeEnum registrationType)
        {
            Guid Id = pluginType.Id ?? Guid.Empty;
            if (Id == Guid.Empty)
            {
                Id = pluginRepository.GetPluginTypeId(parentId, pluginType.Name);
                logWarning?.Invoke($"Extracted id using plugin type name {pluginType.Name}");
            }

            var type = new PluginType()
            {
                Name = pluginType.Name,
                Description = pluginType.Description,
                FriendlyName = pluginType.FriendlyName,
                TypeName = pluginType.TypeName,
                WorkflowActivityGroupName = pluginType.WorkflowActivityGroupName,
                PluginAssemblyId = new EntityReference(PluginAssembly.EntityLogicalName, parentId)
            };

            Id = ExecuteRequest(registrationType, Id, type);
            // AddComponentToSolution(Id, ComponentType.PluginType, solutionName);
            logVerbose?.Invoke($"UpsertPluginType {Id} completed");

            var typeRef = new EntityReference(PluginType.EntityLogicalName, Id);

            foreach (var step in pluginType.Steps)
            {
                var sdkMessageProcessingStepId = UpsertSdkMessageProcessingStep(typeRef, step, solutionName, registrationType);
                logVerbose?.Invoke($"Upsert SdkMessageProcessingStep {sdkMessageProcessingStepId} completed");
                foreach (var image in step.Images)
                {
                    var sdkMessageProcessingStepImageId = UpsertSdkMessageProcessingStepImage(sdkMessageProcessingStepId, image, solutionName, registrationType);
                    logVerbose?.Invoke($"Upsert SdkMessageProcessingStepImage {sdkMessageProcessingStepImageId} completed");
                }
            }
        }

        public List<ServiceEndpt> GetServiceEndpoints(string solutionName, string endPointName) => pluginRepository.GetServiceEndpoints(Guid.Empty, endPointName);

        public void SerializerObjectToFile(string mappingFile, object obj)
        {
            var fileInfo = new FileInfo(mappingFile);
            switch (fileInfo.Extension.ToLower())
            {
                case ".json":
                    Serializers.SaveJson(mappingFile, obj);
                    break;
                case ".xml":
                    Serializers.SaveXml(mappingFile, obj);
                    break;
                default:
                    throw new ArgumentException("Only .json and .xml mapping files are supported", nameof(mappingFile));
            }
        }

        public void UpsertServiceEndpoints(List<ServiceEndpt> serviceEndptLst, string solutionName, RegistrationTypeEnum registrationType)
        {
            foreach (var serviceEndPt in serviceEndptLst)
            {
                logVerbose?.Invoke($"UpsertServiceEndpoint {serviceEndPt.Id} started");
                var serviceEndpointId = UpsertServiceEndpoint(serviceEndPt, solutionName, registrationType);
                logVerbose?.Invoke($"UpsertServiceEndpoint {serviceEndpointId} completed");

                foreach (var step in serviceEndPt.Steps)
                {
                    var serviceEndpointRef = new EntityReference(ServiceEndpoint.EntityLogicalName, serviceEndpointId);
                    logVerbose?.Invoke($"UpsertSdkMessageProcessingStep {step.Id} started");
                    var stepId = UpsertSdkMessageProcessingStep(serviceEndpointRef, step, solutionName, registrationType);
                    logVerbose?.Invoke($"UpsertSdkMessageProcessingStep {stepId} completed");

                    foreach (var image in step.Images)
                    {
                        var stepRef = new EntityReference(SdkMessageProcessingStep.EntityLogicalName, stepId);
                        logVerbose?.Invoke($"UpsertSdkMessageProcessingStepImage {image.Id} started");
                        var imageId = UpsertSdkMessageProcessingStepImage(stepId, image, solutionName, registrationType);
                        logVerbose?.Invoke($"UpsertSdkMessageProcessingStepImage {imageId} completed");
                    }
                }

            }
        }

        private Guid UpsertServiceEndpoint(ServiceEndpt serviceEndpt, string solutionName, RegistrationTypeEnum registrationType)
        {
            Guid Id = serviceEndpt?.Id ?? Guid.Empty;
            if (Id == Guid.Empty)
            {
                Id = pluginRepository.GetServiceEndpointId(serviceEndpt.Name);
                logWarning?.Invoke($"Extracted id using plugin assembly name {serviceEndpt.Name}");
            }

            var serviceEndpoint = new ServiceEndpoint()
            {
                Name = serviceEndpt.Name,
                NamespaceAddress = serviceEndpt.NamespaceAddress,
                ContractEnum = serviceEndpt.Contract,
                Path = serviceEndpt.Path,
                MessageFormatEnum = serviceEndpt.MessageFormat,
                AuthTypeEnum = serviceEndpt.AuthType,
                SASKeyName = serviceEndpt.SASKeyName,
                SASKey = serviceEndpt.SASKey,
                SASToken = serviceEndpt.SASToken,
                UserClaimEnum = serviceEndpt.UserClaim,
                Description = serviceEndpt.Description,
                Url = serviceEndpt.Url,
                AuthValue = serviceEndpt.AuthValue,
            };

            if (!Id.Equals(Guid.Empty) && registrationType == RegistrationTypeEnum.Reset)
            {
                SolutionComponentsManager.DeleteObjectWithDependencies(Id, ComponentType.ServiceEndpoint);
            }

            logVerbose?.Invoke($"Trying to upsert {serviceEndpt.Name} / {Id}");
            Id = ExecuteRequest(registrationType, Id, serviceEndpoint);

            SolutionComponentsManager.AddComponentToSolution(Id, ComponentType.ServiceEndpoint, solutionName);

            return Id;
        }

        public Guid UpsertSdkMessageProcessingStep(EntityReference parentRef, Step step, string solutionName, RegistrationTypeEnum registrationType)
        {
            Guid Id = step.Id ?? Guid.Empty;
            if (Id == Guid.Empty)
            {
                Id = pluginRepository.GetSdkMessageProcessingStepId(parentRef.Id, step.Name);
                logWarning?.Invoke($"Extracted id using plugin step name {step.Name}");
            }

            var sdkMessageId = pluginRepository.GetSdkMessageId(step.MessageName);
            var sdkMessageFilterId = pluginRepository.GetSdkMessageFilterId(step.PrimaryEntityName, sdkMessageId);
            var sdkMessageProcessingStep = new SdkMessageProcessingStep()
            {
                Name = step.Name,
                Description = step.Description,
                SdkMessageId = new EntityReference(SdkMessage.EntityLogicalName, sdkMessageId),
                Configuration = step.CustomConfiguration,
                FilteringAttributes = step.FilteringAttributes,
                ImpersonatingUserId = new EntityReference(SystemUser.EntityLogicalName, pluginRepository.GetUserId(step.ImpersonatingUserFullname)),
                ModeEnum = step.Mode,
                SdkMessageFilterId = sdkMessageFilterId.Equals(Guid.Empty) ? null : new EntityReference(SdkMessageFilter.EntityLogicalName, sdkMessageFilterId),
                Rank = step.Rank,
                StageEnum = step.Stage,
                SupportedDeploymentEnum = step.SupportedDeployment,
                EventHandler = parentRef,
                AsyncAutoDelete = step.AsyncAutoDelete,
            };

            Id = ExecuteRequest(registrationType, Id, sdkMessageProcessingStep);
            int stateCode = (int)step.StateCode;
            organizationService.Execute(new SetStateRequest
            {
                EntityMoniker = new EntityReference(sdkMessageProcessingStep.LogicalName, Id),
                State = new OptionSetValue(stateCode),
                Status = new OptionSetValue(stateCode + 1)
            });

            SolutionComponentsManager.AddComponentToSolution(Id, ComponentType.SDKMessageProcessingStep, solutionName);
            return Id;
        }

        private Guid UpsertSdkMessageProcessingStepImage(Guid parentId, Image image, string solutionName, RegistrationTypeEnum registrationType)
        {
            Guid Id = image.Id ?? Guid.Empty;

            if (Id == Guid.Empty)
            {
                Id = pluginRepository.GetSdkMessageProcessingStepImageId(parentId, image.EntityAlias, image.ImageType);
                logWarning?.Invoke($"Extracted id using plugin step image name {image.EntityAlias}");
            }

            var sdkMessageProcessingStepImage = new SdkMessageProcessingStepImage()
            {
                Attributes1 = image.Attributes,
                EntityAlias = image.EntityAlias,
                MessagePropertyName = image.MessagePropertyName,
                ImageTypeEnum = image.ImageType,
                SdkMessageProcessingStepId = new EntityReference(SdkMessageProcessingStep.EntityLogicalName, parentId)
            };

            Id = ExecuteRequest(registrationType, Id, sdkMessageProcessingStepImage);

            return Id;
        }

        private Guid ExecuteRequest(RegistrationTypeEnum registrationType, Guid Id, Entity entity)
        {
            if (Id != Guid.Empty)
            {
                entity.Id = Id;
            }

            if (registrationType == RegistrationTypeEnum.Upsert)
            {
                entity.Id = Id;
                var query = new QueryExpression(entity.LogicalName) { Criteria = new FilterExpression(), ColumnSet = new ColumnSet(columns: new[] { entity.LogicalName + "id" }) };
                query.Criteria.AddCondition(entity.LogicalName + "id", ConditionOperator.Equal, Id);
                var ids = organizationService.RetrieveMultiple(query);

                if ((ids?.Entities.FirstOrDefault()?.Id ?? Guid.Empty) != Guid.Empty)
                {
                    organizationService.Update(entity);
                }
                else
                {
                    organizationService.Create(entity);
                }
            }
            else
            {
                Id = organizationService.Create(entity);
            }

            return Id;
        }

        

        public object GetPluginRegistrationObject(string assemblyPath, string customAttributeClass)
        {
            reflectionLoader.Initialise(assemblyPath, customAttributeClass);
            return pluginRegistrationObjectFactory.GetAssembly(reflectionLoader);
        }

        public Assembly ReadMappingFile(string mappingFile)
        {
            var fileInfo = new FileInfo(mappingFile);
            switch (fileInfo.Extension.ToLower())
            {
                case ".json":
                    logVerbose("Reading mapping json file");
                    var pluginAssembly = Serializers.ParseJson<Assembly>(mappingFile);
                    logVerbose("Deserialized mapping json file");
                    return pluginAssembly;
                case ".xml":
                    logVerbose("Reading mapping xml file");
                    pluginAssembly = Serializers.ParseXml<Assembly>(mappingFile);
                    logVerbose("Deserialized mapping xml file");
                    return pluginAssembly;
                default:
                    throw new ArgumentException("Only .json and .xml mapping files are supported", nameof(ReadMappingFile));
            }
        }
    }
}