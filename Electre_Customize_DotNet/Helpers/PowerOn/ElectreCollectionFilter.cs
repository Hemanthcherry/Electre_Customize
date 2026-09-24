using Electre_Customize_DotNet.Objects;

namespace Electre_Customize_DotNet.Helpers.PowerOn
{
    internal static class ElectreCollectionFilter
    {
        public static List<ElectreObject> ByPanel(List<ElectreObject> source, IEnumerable<string> selected)
        {
            return ByName(source, selected, e => e.Panel);
        }

        public static List<ElectreObject> BySheet(List<ElectreObject> source, IEnumerable<string> selected)
        {
            return ByName(source, selected, e => e.SheetName);
        }

        public static List<ElectreObject> ByLoom(List<ElectreObject> source, IEnumerable<string> selected)
        {
            return ByName(source, selected, e => e.BundleName);
        }

        private static List<ElectreObject> ByName(
            List<ElectreObject> source,
            IEnumerable<string> selected,
            Func<ElectreObject, string> name)
        {
            var set = new HashSet<string>(selected.Where(s => s != null), StringComparer.OrdinalIgnoreCase);
            var result = new List<ElectreObject>();
            for (int i = 0; i < source.Count; i++)
            {
                var e = source[i];
                string n = name(e);
                if (n != null && set.Contains(n))
                    result.Add(e);
            }
            return result;
        }
    }
}
