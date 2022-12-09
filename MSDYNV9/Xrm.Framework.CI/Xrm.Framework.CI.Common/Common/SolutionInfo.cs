using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xrm.Framework.CI.Common.Entities;

namespace Xrm.Framework.CI.Common
{
    public class SolutionInfo
    {
        public Guid SolutionId { get; set; }
        public string UniqueName { get; set; }


        public static SolutionInfo GetFromSolution(Solution solution) => new SolutionInfo()
        {
            SolutionId = solution.Id,
            UniqueName = solution.UniqueName
        };

        public static SolutionInfo GetFromSolutionId(CIContext context, Guid id) => GetFromSolution((new SolutionManagementRepository(context)).GetEntityById<Solution>(id, x => new Solution()
        {
            UniqueName = x.UniqueName,
            SolutionId = x.SolutionId
        }));

        public override string ToString() => UniqueName;
    }
}
