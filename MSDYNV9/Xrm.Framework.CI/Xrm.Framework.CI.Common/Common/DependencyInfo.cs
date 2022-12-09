using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xrm.Framework.CI.Common.Entities;

namespace Xrm.Framework.CI.Common.Common
{
    /// <summary>
    /// Reasoning about dependency: You cannot delete Required component unless you get rid of Dependent component (Except  DependencyType = Solution Internal). For example attribute Dependency will Require entity it belongs to.
    /// </summary>
    public class DependencyInfo
    {
        public Guid DependencyId { get; set; }
        public DependencyType DependencyType { get; set; }
        public ComponentInfo DependentComponent { get; set; }
        public Guid DependentComponentObjectId { get => DependentComponent.ObjectId; }
        public ComponentType DependentComponentComponentType { get => DependentComponent.ComponentType; }
        public string DependentComponentName { get => DependentComponent.Name; }
        public SolutionInfo DependentComponentBaseSolution { get; set; }
        public Guid DependentComponentBaseSolutionId { get => DependentComponentBaseSolution.SolutionId; }
        public string DependentComponentBaseSolutionUniqueName { get => DependentComponentBaseSolution.UniqueName; }
        public ComponentInfo RequiredComponent { get; set; }
        public Guid RequiredComponentObjectId { get => RequiredComponent.ObjectId; }
        public ComponentType RequiredComponentComponentType { get => RequiredComponent.ComponentType; }
        public string RequiredComponentName { get => RequiredComponent.Name; }
        public SolutionInfo RequiredComponentBaseSolution { get; set; }
        public Guid RequiredComponentBaseSolutionId { get => RequiredComponentBaseSolution.SolutionId; }
        public string RequiredComponentBaseSolutionUniqueName { get => RequiredComponentBaseSolution.UniqueName; }

        DependencyInfo() {}


        public static DependencyInfo GetFromDependency(CIContext context, Dependency dependency)
        {
            var solutionManagementRepository = new SolutionManagementRepository(context);
            var result = new DependencyInfo()
            {
                DependencyId = dependency.DependencyId.Value,
                DependencyType = dependency.DependencyTypeEnum.Value,
                DependentComponent = ComponentInfo.GetFromComponent(context, dependency.DependentComponentObjectId.Value, dependency.DependentComponentTypeEnum.Value),
                DependentComponentBaseSolution = SolutionInfo.GetFromSolutionId(context, dependency.DependentComponentBaseSolutionId.Value),
                RequiredComponent = ComponentInfo.GetFromComponent(context, dependency.RequiredComponentObjectId.Value, dependency.RequiredComponentTypeEnum.Value),
                RequiredComponentBaseSolution = SolutionInfo.GetFromSolutionId(context, dependency.RequiredComponentBaseSolutionId.Value),
            };
            return result;
        }

        public override string ToString() => $"{DependentComponent} -> {RequiredComponent}";
    }
}
