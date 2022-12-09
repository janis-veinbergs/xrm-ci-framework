using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Xrm.Framework.CI.Common.Entities
{
    public static class CIContextExtended
    {
        /// <summary>
        /// Get entity by metadataId - you can use solutioncomponent.objectid for componenttype=1 (entity) to retrieve entity metadata.
        /// </summary>
        /// <param name="service"></param>
        /// <param name="metadataId"></param>
        /// <returns></returns>
        public static EntityMetadata GetEntityMetadata(this CIContext context, Guid metadataId) => ((RetrieveEntityResponse)context.Execute(new RetrieveEntityRequest
        {
            MetadataId = metadataId,
            EntityFilters = EntityFilters.Entity
        })).EntityMetadata;

        public static RelationshipMetadataBase GetRelationshipMetadata(this CIContext context, Guid metadataId) => ((RetrieveRelationshipResponse)context.Execute(new RetrieveRelationshipRequest()
        {
            MetadataId = metadataId,
            RetrieveAsIfPublished = false
        })).RelationshipMetadata;

        public static OptionSetMetadataBase GetOptionSetMetadata(this CIContext context, Guid metadataId) => ((RetrieveOptionSetResponse)context.Execute(new RetrieveOptionSetRequest()
        {
            MetadataId = metadataId
        })).OptionSetMetadata;

        public static AttributeMetadata GetAttributeMetadata(this CIContext context, Guid metadataId) => ((RetrieveAttributeResponse)context.Execute(new RetrieveAttributeRequest()
        {
            MetadataId = metadataId
        })).AttributeMetadata;
    }
}
