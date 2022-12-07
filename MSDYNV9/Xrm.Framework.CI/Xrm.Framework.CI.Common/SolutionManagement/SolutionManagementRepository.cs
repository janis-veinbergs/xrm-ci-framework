using System;
using System.Collections.Concurrent;
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

        public T GetEntityById<T>(Guid id, Func<T, T> selector) where T : Entity => context.CreateQuery<T>().Where(x => x.Id == id).Select(selector).Single();

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

        static ConcurrentDictionary<Guid, Lazy<List<Solution>>> getSolutionsContainingObjectCache = new ConcurrentDictionary<Guid, Lazy<List<Solution>>>(1, 100);
        public IEnumerable<Solution> GetSolutionsContainingObject(Guid objectId) {
            var solutions = getSolutionsContainingObjectCache.GetOrAdd(objectId, new Lazy<List<Solution>>(() => 
                (from c in context.SolutionComponentSet
                 join s in context.SolutionSet on c.SolutionId equals new EntityReference(Solution.EntityLogicalName, s.Id)
                 where c.ObjectId == objectId
                 select new Solution
                 {
                     UniqueName = s.UniqueName
                 }).ToList()
            ));
            return solutions.Value;
        }
    }
}
