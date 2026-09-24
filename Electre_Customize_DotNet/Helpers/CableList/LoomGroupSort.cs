using System.Text.RegularExpressions;

namespace Electre_Customize_DotNet.Helpers.CableList
{
    internal static class LoomGroupSort
    {
        private static readonly Regex GroupStartsWithLetter = new(@"^[A-Za-z]", RegexOptions.Compiled);
        private static readonly Regex GroupMainAndAlpha = new(@"^(\d+)([A-Za-z]*)", RegexOptions.Compiled);
        private static readonly Regex GroupSplitDelims = new(@"[-_]", RegexOptions.Compiled);
        public static readonly Regex ValidGroupPattern = new(@"^[\dA-Za-z_-]+$", RegexOptions.Compiled);

        public static bool GroupStartsWithAlpha(string group)
        {
            return !string.IsNullOrEmpty(group) && GroupStartsWithLetter.IsMatch(group);
        }

        public static List<T> FilterAndSort<T>(List<T> source, Func<T, string> getGroup, Func<T, string> getWireCode)
        {
            int n = source.Count;
            var keys = new List<(T Item, int Main, string Alpha, int Sub, string Wire)>(n);
            for (int i = 0; i < n; i++)
            {
                T item = source[i];
                string group = getGroup(item);
                if (group != null && GroupStartsWithLetter.IsMatch(group))
                    continue;

                string g = group ?? "";
                var match = GroupMainAndAlpha.Match(g);
                int main = match.Success && int.TryParse(match.Groups[1].Value, out int m) ? m : int.MaxValue;
                string alpha = match.Success ? match.Groups[2].Value : "";
                var parts = GroupSplitDelims.Split(g.Replace('_', '-'));
                int sub = parts.Length > 1 && int.TryParse(parts[1], out int subNumber) ? subNumber : 0;
                keys.Add((item, main, alpha, sub, getWireCode(item) ?? ""));
            }

            keys.Sort((a, b) =>
            {
                int c = a.Main.CompareTo(b.Main);
                if (c != 0) return c;
                c = StringComparer.OrdinalIgnoreCase.Compare(a.Alpha, b.Alpha);
                if (c != 0) return c;
                c = a.Sub.CompareTo(b.Sub);
                if (c != 0) return c;
                return string.CompareOrdinal(a.Wire, b.Wire);
            });

            var result = new List<T>(keys.Count);
            for (int i = 0; i < keys.Count; i++)
                result.Add(keys[i].Item);
            return result;
        }
    }
}
