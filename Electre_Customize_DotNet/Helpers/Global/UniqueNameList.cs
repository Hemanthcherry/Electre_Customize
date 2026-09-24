using System.Runtime.CompilerServices;

namespace Electre_Customize_DotNet.Helpers.Global
{
    /// <summary>
    /// First-seen unique names: HashSet lookup plus insertion-order List.
    /// </summary>
    internal static class UniqueNameList
    {
        private static readonly ConditionalWeakTable<List<string>, HashSet<string>> Lookup = new();

        public static bool SearchAndAppend(string istrName, ref List<string> iarr)
        {
            if (istrName == null)
                return false;

            List<string> list = iarr;
            var set = Lookup.GetValue(list, key =>
            {
                var lookup = new HashSet<string>();
                foreach (var item in key)
                {
                    if (item != null)
                        lookup.Add(item);
                }
                return lookup;
            });
            if (set.Count != list.Count)
            {
                set.Clear();
                foreach (var item in list)
                {
                    if (item != null)
                        set.Add(item);
                }
            }

            if (!set.Add(istrName))
                return false;

            list.Add(istrName);
            return true;
        }

        public static bool SearchAndAppend(string istrName, HashSet<string> set)
        {
            if (istrName == null || set == null)
                return false;
            return set.Add(istrName);
        }

        public static bool FastContains(List<string> list, string value)
        {
            if (value == null || list == null || list.Count == 0)
                return false;

            List<string> local = list;
            if (Lookup.TryGetValue(local, out var set) && set.Count == local.Count)
                return set.Contains(value);

            return local.Contains(value);
        }

        public static void Clear(List<string> list)
        {
            if (list == null)
                return;
            list.Clear();
            List<string> local = list;
            if (Lookup.TryGetValue(local, out var set))
                set.Clear();
        }

        public static List<string> ToSortedList(HashSet<string> set)
        {
            var list = new List<string>(set.Count);
            foreach (var item in set)
                list.Add(item);
            list.Sort();
            return list;
        }
    }
}
