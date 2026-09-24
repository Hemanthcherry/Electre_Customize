using System;
using System.Collections.Generic;
using System.Linq;
using Electre_Customize_DotNet.Objects;

namespace Electre_Customize_DotNet.Helpers.PowerOn
{
    internal static class ComponentSelectors
    {
        /// <summary>SCB and TCB components, SCB first, then by panel and connector name.</summary>
        public static List<ElectreObject> ScbTcbSource(List<ElectreObject> panelCollection)
        {
            return panelCollection
                .Where(e => e.ComponentType.Contains("SCB", StringComparison.OrdinalIgnoreCase) ||
                            e.ComponentType.Contains("TCB", StringComparison.OrdinalIgnoreCase))
               .OrderBy(e => e.ComponentType.Contains("SCB", StringComparison.OrdinalIgnoreCase) ? 0 : 1)  // SCB first
                .ThenBy(e => e.Panel)
                  .ThenBy(e => e.ConnectorName)
                 .ToList();
        }
    }
}
