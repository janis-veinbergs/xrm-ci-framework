using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System;
using System.Linq;
using System.Management.Automation;
using Xrm.Framework.CI.Common;
using Xrm.Framework.CI.Common.Entities;

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

        SolutionComponentsManager SolutionComponentsManager;
        #region Process Record

        protected override void BeginProcessing()
        {
            base.BeginProcessing();
            SolutionComponentsManager = new SolutionComponentsManager(Logger, OrganizationService);
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            base.WriteVerbose(string.Format("Removing Solution Components: {0}", SolutionName));

            using (var context = new CIContext(OrganizationService))
            {
                if (Delete.IsPresent)
                {
                    ProcessDelete(SolutionComponent.ObjectId.Value, SolutionComponent.ComponentType);
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

        void ProcessDelete(Guid componentId, OptionSetValue componentType)
        {
            if (!Delete.IsPresent) { throw new InvalidOperationException($"{nameof(ProcessDelete)} not callable without {nameof(Delete)} switch parameter set"); }
            var componentTypeEnum = (ComponentType)componentType.Value;
            if (ShouldProcess($"{componentTypeEnum} {componentId}", $"{this.MyInvocation.MyCommand.Name} -Delete"))
            {
                SolutionComponentsManager.DeleteObjectWithDependencies(componentId, componentTypeEnum);
            }
            base.WriteVerbose($"Deleted component with Id : {componentId} and Type: {componentTypeEnum}");
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
        #endregion
    } 
}