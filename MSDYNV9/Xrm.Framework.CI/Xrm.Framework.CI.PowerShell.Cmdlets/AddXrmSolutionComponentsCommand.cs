using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Xrm.Framework.CI.Common.Common;
using Xrm.Framework.CI.Common.Entities;
using Xrm.Framework.CI.PowerShell.Cmdlets.Common;

namespace Xrm.Framework.CI.PowerShell.Cmdlets
{
    /// <summary>
    /// <para type="synopsis">Copies a CRM Solution Components.</para>
    /// <para type="description">The Move-XrmSolutionComponents of a CRM solution to another by unique name.
    /// </para>
    /// </summary>
    /// <example>
    ///   <code>C:\PS>Move-XrmSolutionComponents -ConnectionString "" -FromSolutionName "UniqueSolutionName -ToSolutionName "UniqueSolutionName"</code>
    ///   <para>Exports the "" managed solution to "" location</para>
    /// </example>
    [Cmdlet(VerbsCommon.Add, "XrmSolutionComponents")]
    public class AddrmSolutionComponentsCommand : XrmCommandBase
    {
        #region Parameters

        /// <summary>
        /// Component to add to solution
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true)]
        public ComponentInfo ComponentInfo { get; set; }

        /// <summary>
        /// To which solution component is to be added
        /// </summary>
        [Parameter(Mandatory = true)]
        public string SolutionName { get; set; }

        #endregion

        #region Process Record

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            base.WriteVerbose($"Adding component to solution {SolutionName}: {ComponentInfo}");


            using (var context = new CIContext(OrganizationService))
            {
                if (ComponentExistsInSolution(context, SolutionName, ComponentInfo.ObjectId, ComponentInfo.ComponentType))
                {
                    return;
                }

                var addReq = new AddSolutionComponentRequest()
                {
                    ComponentId = ComponentInfo.ObjectId,
                    ComponentType = (int)ComponentInfo.ComponentType,
                    AddRequiredComponents = false,
                    SolutionUniqueName = SolutionName
                };

                if (ComponentInfo.IsMetadata)
                {
                    addReq.DoNotIncludeSubcomponents = true;
                    base.WriteVerbose($"DoNotIncludeSubcomponents set to true for {ComponentInfo}");
                }
                OrganizationService.Execute(addReq);
                base.WriteVerbose($"Added component to solution {SolutionName}: {ComponentInfo}");
            }
        }


        bool ComponentExistsInSolution(CIContext context, string solutionName, Guid objectId, ComponentType componentType) => (from sc in context.SolutionComponentSet
                                                                                                                               join s in context.SolutionSet on sc.SolutionId equals new EntityReference(Solution.EntityLogicalName, s.Id)
                                                                                                                               where s.UniqueName == solutionName && sc.ObjectId == objectId && sc.ComponentTypeEnum == componentType
                                                                                                                               select 1).Any();
        #endregion
    }
}