using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Xrm.Framework.CI.Common;
using Xrm.Framework.CI.Common.Common;
using Xrm.Framework.CI.Common.Entities;
using Xrm.Framework.CI.PowerShell.Cmdlets.Common;

namespace Xrm.Framework.CI.PowerShell.Cmdlets
{
    /// <summary>
    /// <para type="synopsis">Gets CRM Solution Components.</para>
    /// <para type="description"></para>
    /// </summary>
    /// <example>
    ///   <code>C:\PS>Get-XrmSolutionComponents -ConnectionString "" -SolutionName "UniqueSolutionName"</code>
    ///   <para>Gets UniqueSolutionName components</para>
    /// </example>
    [Cmdlet(VerbsCommon.Get, "XrmSolutionComponentsForDelete")]
    [OutputType(typeof(SolutionComponentInfo))]
    public class GetXrmSolutionComponentsForDeleteCommand : XrmCommandBase
    {
        #region Parameters

        /// <summary>
        /// <para type="description">The unique name of the solution to get component froms.</para>
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "SolutionComponentInfo")]
        public SolutionComponentInfo SolutionComponentInfo { get; set; }

        /// <summary>
        /// <para type="description">The unique name of the solution to get component froms.</para>
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "SolutionComponent")]
        public SolutionComponent SolutionComponent { get; set; }

        SolutionManagementRepository solutionManagementRepository;
        #endregion

        #region Process Record

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            base.WriteVerbose($"Getting Solution Components for Delete from {SolutionComponentInfo?.Name ?? SolutionComponent?.Id.ToString()}");

            using (var context = new CIContext(OrganizationService))
            {
                if (solutionManagementRepository == null)
                {
                    solutionManagementRepository = new SolutionManagementRepository(context);
                }

                var solutionComponentEntity = SolutionComponent ?? solutionManagementRepository.GetEntityById<SolutionComponent>(SolutionComponentInfo.SolutionComponentId);
                foreach (var x in GetComponentsForDelete(context, solutionComponentEntity))
                {
                    WriteObject(x);
                }
            }
        }

        public IEnumerable<SolutionComponentInfo> GetComponentsForDelete(CIContext context, SolutionComponent solutionComponent, int depth = 0, Guid? parentComponentId = null, HashSet<string> processedHashSet = null)
        {
            if (processedHashSet == null)
            {
                processedHashSet = new HashSet<string>();
            }
            var objectkey = $"{solutionComponent.ComponentTypeEnum}{solutionComponent.Id}";

            var componentInfo = SolutionComponentInfo.GetFromComponent(context, OrganizationService, solutionComponent, null, depth, parentComponentId);
            yield return componentInfo;
            if (processedHashSet.Contains(objectkey))
            {
                yield break;
            }
            processedHashSet.Add(objectkey);

            foreach (var dependency in solutionManagementRepository.GetDependeciesForDelete(solutionComponent.ObjectId.Value, solutionComponent.ComponentTypeEnum))
            {
                SolutionComponent solutionComponentDependency = null;
                try
                {
                    solutionComponentDependency = context.CreateQuery<SolutionComponent>().Where(x => x.ObjectId == dependency.DependentComponentObjectId.Value && x.SolutionId == new EntityReference(Solution.EntityLogicalName, dependency.DependentComponentBaseSolutionId.Value)).FirstOrDefault();
                }
                catch (InvalidOperationException)
                {
                    //meh
                }
                if (solutionComponentDependency != null)
                {
                    foreach (var x in GetComponentsForDelete(context, solutionComponentDependency, depth + 1, solutionComponent.SolutionComponentId, processedHashSet))
                    {
                        yield return x;
                    }
                }
            }
        }

        #endregion
    }
}