using Electre_Customize_DotNet.Objects;

namespace Electre_Customize_DotNet.Helpers.PowerOn
{
    internal sealed class GroundPinIndex
    {
        private readonly Dictionary<string, List<ElectreObject>> _byPanel;

        private GroundPinIndex(Dictionary<string, List<ElectreObject>> byPanel)
        {
            _byPanel = byPanel;
        }

        public static GroundPinIndex Build(List<ElectreObject> elecCollection)
        {
            var byPanel = new Dictionary<string, List<ElectreObject>>();
            if (elecCollection == null)
                return new GroundPinIndex(byPanel);

            for (int i = 0; i < elecCollection.Count; i++)
            {
                var e = elecCollection[i];
                if (!IsGroundType(e.ComponentType))
                    continue;
                string panel = e.Panel ?? "";
                if (!byPanel.TryGetValue(panel, out var list))
                {
                    list = new List<ElectreObject>();
                    byPanel[panel] = list;
                }
                list.Add(e);
            }
            return new GroundPinIndex(byPanel);
        }

        public List<ElectreObject> FindForSource(ElectreObject source)
        {
            var result = new List<ElectreObject>();
            if (source == null)
                return result;
            if (!_byPanel.TryGetValue(source.Panel ?? "", out var list))
                return result;

            string prefix = (source.ConnectorName ?? "") + "_RTN";
            for (int i = 0; i < list.Count; i++)
            {
                var e = list[i];
                if (e.ConnectorName != null
                    && e.ConnectorName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    result.Add(e);
            }
            return result;
        }

        private static bool IsGroundType(string componentType)
        {
            if (string.IsNullOrEmpty(componentType))
                return false;
            return componentType.Contains("GROUND", StringComparison.OrdinalIgnoreCase)
                || componentType.Contains("TER", StringComparison.OrdinalIgnoreCase)
                || componentType.Contains("TBK", StringComparison.OrdinalIgnoreCase);
        }
    }
}
