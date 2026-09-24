using Electre_Customize_DotNet.Helpers.Global;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace Electre_Customize_DotNet.Helpers.CableList
{
    internal readonly struct HalMeMerge
    {
        public readonly string From;
        public readonly string To;
        public readonly string KTo;
        public readonly object Serial;
        public readonly bool ClearK;

        public HalMeMerge(string from, string to, string kTo, object serial, bool clearK)
        {
            From = from;
            To = to;
            KTo = kTo;
            Serial = serial;
            ClearK = clearK;
        }
    }

    internal static class CableListMerges
    {
        public static void MergeMe(
            Worksheet ws,
            object iFrom,
            object iTo,
            object serial,
            object classValue,
            object shieldExtra,
            object cableClassification,
            object lengthValue)
        {
            ApplyMeBatch(ws, new[]
            {
                BuildMe(iFrom, iTo, serial, shieldExtra, cableClassification)
            });
        }

        public static HalMeMerge BuildMe(object iFrom, object iTo, object serial, object shieldExtra, object cableClassification)
        {
            string from = Convert.ToString(iFrom) ?? "";
            string to = Convert.ToString(iTo) ?? "";
            int extra = 0;
            try { extra = Convert.ToInt32(shieldExtra); } catch { extra = 0; }

            string kTo = to;
            bool clearK = false;
            if (extra != 0 && string.Equals(Convert.ToString(cableClassification), "Special", StringComparison.OrdinalIgnoreCase))
            {
                if (int.TryParse(to, out int toNum))
                    kTo = (toNum + extra).ToString();
                clearK = true;
            }
            return new HalMeMerge(from, to, kTo, serial, clearK);
        }

        public static void ApplyMeBatch(Worksheet ws, IList<HalMeMerge> ops)
        {
            if (ops == null || ops.Count == 0)
                return;

            var bAreas = new List<string>(ops.Count);
            var dAreas = new List<string>(ops.Count);
            var kAreas = new List<string>(ops.Count);
            for (int i = 0; i < ops.Count; i++)
            {
                HalMeMerge op = ops[i];
                if (op.From != op.To)
                {
                    bAreas.Add("B" + op.From + ":B" + op.To);
                    dAreas.Add("D" + op.From + ":D" + op.To);
                }
                if (op.From != op.KTo)
                    kAreas.Add("K" + op.From + ":K" + op.KTo);
            }

            MergeAreaList(ws, bAreas);
            MergeAreaList(ws, dAreas);
            MergeAreaList(ws, kAreas);

            for (int i = 0; i < ops.Count; i++)
            {
                HalMeMerge op = ops[i];
                ws.Range["B" + op.From].Value2 = op.Serial;
                if (op.ClearK)
                    ws.Range["K" + op.From].Value2 = "";
            }
        }

        private static void MergeAreaList(Worksheet ws, List<string> areas)
        {
            const int batch = 8;
            for (int i = 0; i < areas.Count; i += batch)
            {
                int n = Math.Min(batch, areas.Count - i);
                string addr = areas[i];
                for (int j = 1; j < n; j++)
                    addr = addr + "," + areas[i + j];
                Range rng = ws.Range[addr];
                rng.Merge();
                ExcelRangeHelper.ReleaseCom(rng);
            }
        }

        public static void UnmergeRow(Worksheet ws, object iFrom, object serial, object classValue, object cableClassification)
        {
            string from = Convert.ToString(iFrom) ?? "";

            Range rangeB = ws.Range["B" + from, "B" + from];
            rangeB.MergeCells = false;
            rangeB.WrapText = true;
            rangeB.Value2 = serial;
            ExcelRangeHelper.ReleaseCom(rangeB);

            Range rangeK = ws.Range["K" + from, "K" + from];
            if (string.Equals(Convert.ToString(cableClassification), "Normal", StringComparison.OrdinalIgnoreCase))
                rangeK.Value2 = classValue;
            else if (string.Equals(Convert.ToString(cableClassification), "Special", StringComparison.OrdinalIgnoreCase))
                rangeK.Value2 = "";
            ExcelRangeHelper.ReleaseCom(rangeK);
        }
    }
}
