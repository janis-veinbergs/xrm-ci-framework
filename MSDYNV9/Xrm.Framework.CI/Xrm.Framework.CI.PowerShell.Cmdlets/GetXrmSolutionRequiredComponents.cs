using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.ServiceModel;
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
    /// <para type="synopsis">Gets CRM Solution Components that another solution component depends on.</para>
    /// <para type="description">This is reverse of Get-XrmSolutionDependentComponents</para>
    /// </summary>
    /// <example>
    ///   <code>C:\PS>Get-XrmSolutionComponents -SolutionName "UniqueSolutionName" -ConnectionString "" | Get-XrmSolutionRequiredComponents -ConnectionString ""</code>
    ///   <para>Returns a list of the dependencies for solution components that another solution component directly depends on</para>
    /// </example>
    [Cmdlet(VerbsCommon.Get, "XrmSolutionRequiredComponents")]
    [OutputType(typeof(DependencyInfo))]
    public class GetXrmSolutionRequiredComponentsCommand : XrmCommandBase
    {
        #region Parameters

        /// <summary>
        /// <para type="description">The unique name of the solution to get component froms.</para>
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "ComponentInfo")]
        public ComponentInfo ComponentInfo { get; set; }

        /// <summary>
        /// <para type="description">The unique name of the solution to get component froms.</para>
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "SolutionComponent")]
        public SolutionComponent SolutionComponent { get; set; }

        /// <summary>
        /// Get required components recursively
        /// </summary>
        [Parameter]
        public SwitchParameter Recursive { get; set; }

        /// <summary>
        /// Maximum depth when going recursively in. Used with Recursive
        /// </summary>
        [Parameter]
        public int MaxDepth { get; set; } = int.MaxValue;

        CIContext context;
        SolutionManagementRepository solutionManagementRepository;
        #endregion

        #region Process Record

        protected override void BeginProcessing()
        {
            base.BeginProcessing();
            context = new CIContext(OrganizationService);
            solutionManagementRepository = new SolutionManagementRepository(context);
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            base.WriteVerbose($"Getting Solution Required Components from {ComponentInfo?.Name ?? SolutionComponent?.Id.ToString()}");

            var objectId = SolutionComponent?.ObjectId ?? ComponentInfo.ObjectId;
            var componentType = SolutionComponent?.ComponentTypeEnum ?? ComponentInfo.ComponentType;
            foreach (var x in GetRequiredComponents(context, objectId, componentType))
            {
                WriteObject(x);
            }
        }

        protected override void EndProcessing()
        {
            base.EndProcessing();
            if (context != null)
            {
                context.Dispose();
                context = null;
            }
        }

        IEnumerable<DependencyInfo> GetRequiredComponents(CIContext context, Guid objectId, ComponentType componentType)
        {
            foreach (var dep in GetRequiredComponentsOrWarn(objectId, componentType))
            {
                foreach (var item in GetRequiredComponentsRecursive(context, dep))
                {
                    yield return item;
                }

            }
            yield break;
        }

        IEnumerable<DependencyInfo> GetRequiredComponentsRecursive(CIContext context, Dependency dependency, HashSet<string> processedHashSet = null, int depth = 1)
        {
            if (processedHashSet == null)
            {
                processedHashSet = new HashSet<string>();
            }
            var objectkey = $"{dependency.DependentComponentTypeEnum.Value}{dependency.DependentComponentObjectId.Value}";

            var dependencyInfo = DependencyInfo.GetFromDependency(context, dependency);
            if (dependencyInfo.DependentComponent.Name == null)
            {
                WriteWarning($"Object {dependencyInfo.DependentComponent.ComponentType} {dependencyInfo.DependentComponent.ObjectId} found as a dependant component, but not retrievable, that is entity doesn't exist.");
            }
            if (dependencyInfo.RequiredComponent.Name == null)
            {
                WriteWarning($"Object {dependencyInfo.RequiredComponent.ComponentType} {dependencyInfo.RequiredComponent.ObjectId} found as a required component, but not retrievable, that is entity doesn't exist.");
            }
            yield return dependencyInfo;
            if (processedHashSet.Contains(objectkey) || !Recursive.IsPresent || depth > MaxDepth)
            {
                yield break;
            }
            processedHashSet.Add(objectkey);

            foreach (var childDependency in GetRequiredComponentsOrWarn(dependency.RequiredComponentObjectId.Value, dependency.RequiredComponentTypeEnum.Value))
            {
                foreach (var x in GetRequiredComponentsRecursive(context, childDependency, processedHashSet, depth + 1))
                {
                    yield return x;
                }
            }
        }

        IEnumerable<Dependency> GetRequiredComponentsOrWarn(Guid objectId, ComponentType componentType)
        {
            IEnumerable<Dependency> dependencies = null;
            try
            {
                dependencies = solutionManagementRepository.GetRequiredComponents(objectId, componentType);
            }
            catch (FaultException<OrganizationServiceFault> ex) when (ex.Detail.ErrorDetails.TryGetValue("ApiOriginalExceptionKey", out object originalException) && (originalException as string).StartsWith("Microsoft.Crm.BusinessEntities.CrmObjectNotFoundException"))
            {
                WriteWarning($"Fault when retrieving required components for component {componentType} {objectId}: {ex.Message}");
            }
            if (dependencies != null)
            {
                foreach (var x in dependencies)
                {
                    yield return x;
                }
            }
            yield break;
        }
        #endregion
    }
}