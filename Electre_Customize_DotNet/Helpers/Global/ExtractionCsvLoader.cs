using Electre_Customize_DotNet.Objects;
using System.Text;

namespace Electre_Customize_DotNet.Helpers.Global
{
    internal static class ExtractionCsvLoader
    {
        public static void Load(
            List<ElectreObject> general,
            List<ElectreObject> panel,
            HashSet<string> partNumbers,
            Dictionary<string, HashSet<string>> partToConnectors,
            Dictionary<string, List<string>> connectorToParts)
        {
            general.Clear();
            panel.Clear();

            EnsureListCapacity(general, GlobalVar.DataExtractionGlobal);
            EnsureListCapacity(panel, GlobalVar.DataExtractionGlobal);

            using (StreamReader reader = OpenExtractionReader())
            {
                while (!reader.EndOfStream)
                {
                    string line = reader.ReadLine();
                    if (line == null)
                        break;

                    string[] arrTemp = line.Split(';');
                    CollectPartNumberMaps(arrTemp, partNumbers, partToConnectors, connectorToParts);

                    if (string.IsNullOrEmpty(arrTemp[12]))
                        continue;

                    if (string.IsNullOrEmpty(arrTemp[28]))
                        general.Add(CreateElectreObject(arrTemp));
                    else
                        panel.Add(CreateElectreObject(arrTemp));
                }
            }
        }

        private static StreamReader OpenExtractionReader()
        {
            var fs = new FileStream(
                GlobalVar.DataExtractionGlobal,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read,
                65536,
                FileOptions.SequentialScan);
            return new StreamReader(fs, Encoding.UTF8, true, 65536);
        }

        private static void EnsureListCapacity(List<ElectreObject> list, string path)
        {
            try
            {
                long len = new FileInfo(path).Length;
                int estimate = (int)Math.Clamp(len / 80L, 256, 2_000_000);
                if (list.Capacity < estimate)
                    list.Capacity = estimate;
            }
            catch
            {
                if (list.Capacity < 1024)
                    list.Capacity = 1024;
            }
        }

        private static ElectreObject CreateElectreObject(string[] arrTemp)
        {
            var electreObj = new ElectreObject
            {
                SheetName = arrTemp[0],
                SheetNumber = arrTemp[1],
                DrawingNumber = arrTemp[2],
                DefaultGauge = arrTemp[3],
                BundleName = arrTemp[4],
                EquipmentName = arrTemp[5],
                ConnectorName = arrTemp[6],
                PinNumber = arrTemp[7],
                Ends = arrTemp[8],
                FunctionalDesignation = arrTemp[9],
                SymbolName = arrTemp[10],
                ComponentType = arrTemp[11],
                WireNumber = arrTemp[12],
                Group = arrTemp[13],
                Gauge = arrTemp[14],
                CableType = arrTemp[15],
                Length = arrTemp[16],
                Signal = arrTemp[17],
                Layer = arrTemp[18],
                OverShield = arrTemp[19],
                Net = arrTemp[21],
                SubNet = arrTemp[22],
                Shunt = arrTemp[23],
                Core_Part_Number = arrTemp[24],
                Voltage = arrTemp[27],
                Tag6_Link = "",
                Tag7 = arrTemp[15],
                Panel = arrTemp[28],
                ShuntExt1 = arrTemp[23],
                NoMegger = arrTemp[29]
            };

            if (!string.IsNullOrEmpty(arrTemp[29]))
                electreObj.CableType = arrTemp[29];

            return electreObj;
        }

        private static void CollectPartNumberMaps(
            string[] arrTemp,
            HashSet<string> partNumbers,
            Dictionary<string, HashSet<string>> partToConnectors,
            Dictionary<string, List<string>> connectorToParts)
        {
            string part = arrTemp[24];
            if (string.IsNullOrEmpty(part) || part.Length <= 2)
                return;

            partNumbers.Add(part);

            string connector = arrTemp[6];
            if (string.IsNullOrEmpty(connector))
                return;

            if (!partToConnectors.TryGetValue(part, out var connectors))
            {
                connectors = new HashSet<string>();
                partToConnectors[part] = connectors;
            }
            connectors.Add(connector);

            if (!connectorToParts.TryGetValue(connector, out var parts))
            {
                parts = new List<string>();
                connectorToParts[connector] = parts;
            }
            parts.Add(part);
        }
    }
}
