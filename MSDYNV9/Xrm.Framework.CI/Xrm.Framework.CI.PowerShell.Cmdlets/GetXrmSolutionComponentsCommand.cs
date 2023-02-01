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
    [Cmdlet(VerbsCommon.Get, "XrmSolutionComponents", DefaultParameterSetName = "SolutionComponent")]
    [OutputType(typeof(SolutionComponent), ParameterSetName = new[] { "SolutionComponent" })]
    [OutputType(typeof(ComponentInfo), ParameterSetName = new[] { "ComponentInfo" })]
    public class GetXrmSolutionComponentsCommand : XrmCommandBase
    {
        #region Parameters

        /// <summary>
        /// <para type="description">The unique name of the solution to get component froms.</para>
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = "SolutionComponent")]
        [Parameter(Mandatory = true, ParameterSetName = "ComponentInfo")]
        public string SolutionName { get; set; }

        /// <summary>
        /// Processes only top-level unmanaged solution components. However returned delete dependencies are also returned managed ones.
        /// </summary>
        [Parameter(ParameterSetName = "SolutionComponent")]
        [Parameter(ParameterSetName = "ComponentInfo")]
        public SwitchParameter Unmanaged { get; set; }

        /// <summary>
        /// Get result as <see cref="ComponentInfo"/>. Without this switch, it is outputted as <see cref="SolutionComponent"/> entity.
        /// </summary>
        [Parameter(ParameterSetName = "ComponentInfo")]
        public SwitchParameter AsComponentInfo { get; set; }

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
            base.WriteVerbose(string.Format("Getting Solution Components from {0}", SolutionName));

            var solutionId = GetSolutionId(context, SolutionName);
            var solutionComponents = (from s in context.SolutionComponentSet
                                        where s.SolutionId == new EntityReference(Solution.EntityLogicalName, solutionId)
                                        orderby s.RootSolutionComponentId descending
                                        select s).ToList();

            foreach (var solutionComponent in solutionComponents)
            {
                if (AsComponentInfo.IsPresent)
                {
                    var componentInfo = ComponentInfo.GetFromComponent(context, solutionComponent.ObjectId.Value, solutionComponent.ComponentTypeEnum.Value);
                    WriteObject(componentInfo);
                } else
                {
                    WriteObject(solutionComponent);
                }
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

        private Guid GetSolutionId(CIContext context, string solutionName)
        {
            var query1 = from solution in context.SolutionSet
                         where solution.UniqueName == solutionName
                         select solution.Id;

            if (query1 == null)
            {
                throw new Exception(string.Format("Solution {0} could not be found", solutionName));
            }

            var solutionId = query1.FirstOrDefault();

            return solutionId;
        }

        #endregion
    }
}