using System.Collections.Generic;
using Electre_Customize_DotNet.Objects;

namespace Electre_Customize_DotNet.Helpers.CableList
{
    /// <summary>
    /// Groups the indices of an ElectreObject list by (WireNumber, SubNet), compared case-insensitively
    /// via ToUpper() - exactly the comparison the original per-item full-collection rescans used
    /// (ElecCollection[j].WireNumber.ToUpper() == E1.WireNumber.ToUpper() &amp;&amp; same for SubNet).
    /// Only objects the original loops could ever match are indexed: ComponentType != "SDS" and a
    /// non-empty WireNumber. Indices keep their original ascending order, so iterating a group
    /// reproduces the original "scan j = 0..n-1" order.
    /// </summary>
    internal static class WireSubNetIndex
    {
        /// <summary>The lookup key for an object (throws on a null SubNet, as the original did).</summary>
        public static (string Wire, string SubNet) KeyOf(ElectreObject item)
            => (item.WireNumber.ToUpper(), item.SubNet.ToUpper());

        /// <summary>
        /// Persistent groups: every member can enumerate the whole group. Used where each item must see
        /// all of its partners (Removing_DuplicateWires_In2DArray emits both directions of every pair).
        /// </summary>
        public static Dictionary<(string Wire, string SubNet), List<int>> BuildGroups(IReadOnlyList<ElectreObject> items)
        {
            var groups = new Dictionary<(string Wire, string SubNet), List<int>>();
            for (int k = 0; k < items.Count; k++)
            {
                var candidate = items[k];
                if (candidate.ComponentType != "SDS" && !string.IsNullOrEmpty(candidate.WireNumber))
                {
                    var key = KeyOf(candidate);
                    if (!groups.TryGetValue(key, out var group))
                    {
                        group = new List<int>();
                        groups[key] = group;
                    }
                    group.Add(k);
                }
            }
            return groups;
        }

        /// <summary>
        /// Consumable queues: each index is dequeued at most once across the whole run. Used where a
        /// partner is "used up" once linked (Reading_And_StoringData_In2DArray marks Tag6_Link).
        /// </summary>
        public static Dictionary<(string Wire, string SubNet), Queue<int>> BuildQueues(IReadOnlyList<ElectreObject> items)
        {
            var groups = BuildGroups(items);
            var queues = new Dictionary<(string Wire, string SubNet), Queue<int>>(groups.Count);
            foreach (var pair in groups)
            {
                queues[pair.Key] = new Queue<int>(pair.Value);
            }
            return queues;
        }
    }
}
