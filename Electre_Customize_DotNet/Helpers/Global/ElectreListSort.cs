using Electre_Customize_DotNet.Objects;

namespace Electre_Customize_DotNet.Helpers.Global
{
    internal static class ElectreListSort
    {
        public static void StableByConnectorAndPin(List<ElectreObject> list)
        {
            int n = list.Count;
            if (n < 2)
                return;

            var keyed = new (string Connector, string Pin, int Index, ElectreObject Obj)[n];
            for (int i = 0; i < n; i++)
            {
                var obj = list[i];
                keyed[i] = (obj.ConnectorName, obj.PinNumber, i, obj);
            }

            Array.Sort(keyed, (a, b) =>
            {
                int cmp = string.Compare(a.Connector, b.Connector);
                if (cmp != 0)
                    return cmp;
                cmp = string.Compare(a.Pin, b.Pin);
                if (cmp != 0)
                    return cmp;
                return a.Index.CompareTo(b.Index);
            });

            for (int i = 0; i < n; i++)
                list[i] = keyed[i].Obj;
        }
    }
}
