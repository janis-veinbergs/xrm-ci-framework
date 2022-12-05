using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Xrm.Framework.CI.Common.Entities;

namespace Xrm.Framework.CI.Common
{
    public class SolutionManagementRepository
    {
        private CIContext context { get; }
        private PluginRepository pluginRepository { get; }

        public SolutionManagementRepository(CIContext context)
        {
            this.context = context;
            pluginRepository = new PluginRepository(context);
        }

        public T GetEntityById<T>(Guid id) where T : Entity => context.CreateQuery<T>().Single(x => x.Id == id);

        public Entity GetEntityByReference(EntityReference entityReference) => context.CreateQuery(entityReference.LogicalName).Single(x => x.Id == entityReference.Id);

        public OptionSetMetadataBase GetOptionSetMetadata(Guid metadataId) => ((RetrieveOptionSetResponse)context.Execute(new RetrieveOptionSetRequest()
        {
            MetadataId = metadataId
        })).OptionSetMetadata;


        public SdkMessageProcessingStep GetSdkMessageProcessingStepById(Guid id) => pluginRepository.GetSdkMessageProcessingStepById(id);
        public PluginType GetPluginTypeById(Guid id) => pluginRepository.GetPluginTypeById(id);

        public IEnumerable<Dependency> GetDependeciesForDelete(Guid objectId, ComponentType? componentType) => ((RetrieveDependenciesForDeleteResponse)context.Execute(new RetrieveDependenciesForDeleteRequest()
        {
            ComponentType = (int)componentType,
            ObjectId = objectId
        })).EntityCollection.Entities.Select(x => x.ToEntity<Dependency>());
    }
}
