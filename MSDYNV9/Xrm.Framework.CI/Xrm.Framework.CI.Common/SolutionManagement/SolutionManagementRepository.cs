using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Xrm.Framework.CI.Common.Entities;

namespace Xrm.Framework.CI.Common
{
    public class SolutionManagementRepository
    {
        private CIContext context { get; }

        public SolutionManagementRepository(CIContext context)
        {
            this.context = context;
        }

        public T GetEntityById<T>(Guid id) where T : Entity => context.CreateQuery<T>().Single(x => x.Id == id);

        public OptionSetMetadataBase GetOptionSetMetadata(Guid metadataId) => ((RetrieveOptionSetResponse)context.Execute(new RetrieveOptionSetRequest()
        {
            MetadataId = metadataId
        })).OptionSetMetadata;
    }
}
