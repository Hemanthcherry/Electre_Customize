using System;
using System.Collections.Generic;
using System.Linq;
using Electre_Customize_DotNet.Objects;

namespace Electre_Customize_DotNet.Helpers.PowerOn
{
    /// <summary>
    /// Candidate ground / return components (ComponentType contains GROUND, TER or TBK) grouped by panel,
    /// so PowerOnReport.FindGroundPin does not rescan the whole collection once per circuit breaker.
    ///
    /// The original per-source filter was
    ///   (type is GROUND/TER/TBK) &amp;&amp; e.Panel == source.Panel &amp;&amp; e.ConnectorName.StartsWith(source.ConnectorName + "_RTN", OrdinalIgnoreCase)
    /// evaluated in list order. This keeps the same tests, the same order of results and the same set of
    /// elements that ever reach the StartsWith (so a null ConnectorName behaves as before).
    /// </summary>
    internal sealed class GroundCandidateIndex
    {
        // ValueTuple keys tolerate a null Panel, which a plain string key in a Dictionary would not.
        private readonly Dictionary<ValueTuple<string>, List<ElectreObject>> _byPanel = new Dictionary<ValueTuple<string>, List<ElectreObject>>();

        public GroundCandidateIndex(List<ElectreObject> elecCollection)
        {
            foreach (var e in elecCollection)
            {
                if (e.ComponentType.Contains("GROUND", StringComparison.OrdinalIgnoreCase) ||
                    e.ComponentType.Contains("TER", StringComparison.OrdinalIgnoreCase) ||
                    e.ComponentType.Contains("TBK", StringComparison.OrdinalIgnoreCase))
                {
                    var key = ValueTuple.Create(e.Panel);
                    if (!_byPanel.TryGetValue(key, out var list))
                    {
                        list = new List<ElectreObject>();
                        _byPanel[key] = list;
                    }
                    list.Add(e);
                }
            }
        }

        /// <summary>The ground components of the source's panel whose connector name starts with "{source connector}_RTN".</summary>
        public List<ElectreObject> CandidatesFor(ElectreObject source)
        {
            if (!_byPanel.TryGetValue(ValueTuple.Create(source.Panel), out var list))
            {
                return new List<ElectreObject>();
            }

            string prefix = $"{source.ConnectorName}_RTN";
            return list.Where(e => e.ConnectorName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)).ToList();
        }
    }
}
