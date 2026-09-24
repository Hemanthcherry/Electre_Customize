using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace Electre_Customize_DotNet.Helpers.Global
{
    /// <summary>
    /// "Add to a List&lt;string&gt; only if it is not already there", keeping the list's first-seen order.
    ///
    /// This is called hundreds of thousands of times across a full data load (every row's connector,
    /// bundle, sheet and panel) and again per traced connection in Continuity report generation, always
    /// against lists that keep growing. A plain linear "contains" scan before every append makes each of
    /// those call sites effectively O(n^2). A HashSet mirror is cached per list instance (by reference,
    /// released automatically when the list is garbage collected) so the membership check is O(1); the
    /// List itself and its element order are untouched.
    /// </summary>
    internal static class UniqueStringList
    {
        private static readonly ConditionalWeakTable<List<string>, HashSet<string>> Mirrors = new();

        /// <summary>Appends <paramref name="value"/> unless present. Returns true if it was added.</summary>
        public static bool TryAdd(string value, ref List<string> list)
        {
            var set = SyncedMirror(list);

            if (set.Contains(value))
            {
                return false;
            }

            list.Add(value);
            set.Add(value);
            return true;
        }

        /// <summary>Same answer as <c>list.Contains(value)</c> (default string equality) but O(1) instead of a linear scan.</summary>
        public static bool Contains(string value, List<string> list)
        {
            return SyncedMirror(list).Contains(value);
        }

        private static HashSet<string> SyncedMirror(List<string> list)
        {
            var set = Mirrors.GetValue(list, l => new HashSet<string>(l));

            // Self-heal if the list was mutated directly (Clear(), a direct Add(), or the reference being
            // reused for different contents) without going through this class, so the mirror can never
            // silently drift out of sync with the list it mirrors.
            if (set.Count != list.Count)
            {
                set.Clear();
                foreach (var item in list)
                {
                    set.Add(item);
                }
            }

            return set;
        }
    }
}
