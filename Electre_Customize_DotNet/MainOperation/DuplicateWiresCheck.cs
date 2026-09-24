using Electre_Customize_DotNet.Logs;
using Electre_Customize_DotNet.Objects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Electre_Customize_DotNet.MainOperation
{
    public class DuplicateWiresCheck
    {
        public static void DuplicateWires()
        {
            // Finding duplicate wires in the project
            List<WireListCount> wireListCountList = modMain.ElecCollection_All
            .Where(obj => !string.Equals(obj.PinNumber, "SHLD", StringComparison.OrdinalIgnoreCase))
            .GroupBy(p => new { p.WireNumber, p.Core_Part_Number })
            .Where(g => g.Count() > 2)
            .Select(g => new WireListCount()
            {
                WireNumber = g.Key.WireNumber,
                Core_Part_Number = g.Key.Core_Part_Number,
                WireCount = g.Count(),
            })
            .ToList();

            var monoWireSet = new HashSet<string>(
                modMain.ElecCollection_All
                .Where(e => string.IsNullOrEmpty(e.Core_Part_Number) || e.Core_Part_Number.Length > 2 || string.Equals(e.Core_Part_Number, "PN", StringComparison.OrdinalIgnoreCase))
                .Select(e => e.WireNumber),
                StringComparer.OrdinalIgnoreCase
            );

            //var monoPairConflicts = ElecCollection
            //        .Where(w => !string.IsNullOrEmpty(w.Core_Part_Number)
            //                 && monoWireSet.Contains(w.WireNumber)) 
            //        .ToList();

            var monoPairConflicts = modMain.ElecCollection_All
                .Where(w => !string.IsNullOrEmpty(w.Core_Part_Number) && w.Core_Part_Number.Length <=2 && !string.Equals(w.Core_Part_Number, "PN", StringComparison.OrdinalIgnoreCase) && monoWireSet.Contains(w.WireNumber))
                .GroupBy(w => new { w.WireNumber, w.Core_Part_Number })
                .Select(g => new
                {
                    g.Key.WireNumber,
                    g.Key.Core_Part_Number,
                    WireCount = 3//g.Count()
                })
                .ToList();


            if (monoPairConflicts.Count > 0)
            {
                foreach (var conflict in monoPairConflicts)
                {
                    if (!wireListCountList.Any(w =>
                        string.Equals(w.WireNumber, conflict.WireNumber, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(w.Core_Part_Number, conflict.Core_Part_Number, StringComparison.OrdinalIgnoreCase)))
                    {
                        wireListCountList.Add(new WireListCount
                        {
                            WireNumber = conflict.WireNumber,
                            Core_Part_Number = conflict.Core_Part_Number,
                            WireCount = conflict.WireCount
                        });
                    }
                }
            }

            if (wireListCountList.Any(e => e.WireCount > 2))
            {
                List<DuplicateWire> duplicateWires = new List<DuplicateWire>();

                // Convert wireListCountList to a HashSet for faster lookup
                //var wireSet = new HashSet<(string, string)>(
                //    wireListCountList.Select(w => (w.WireNumber, w.Core_Part_Number))
                //);
                var wireSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var w in wireListCountList)
                    wireSet.Add((w.WireNumber ?? "") + "\u001f" + (w.Core_Part_Number ?? ""));

                foreach (ElectreObject elec in modMain.ElecCollection_All)
                {
                    string key = (elec.WireNumber ?? "") + "\u001f" + (elec.Core_Part_Number ?? "");

                    if (wireSet.Contains(key))
                    {
                        duplicateWires.Add(new DuplicateWire
                        {
                            WireName = elec.WireNumber,
                            WireNumber = elec.Core_Part_Number ?? string.Empty,
                            sheetName = elec.SheetName
                        });
                    }
                }

                var wireDuplicateFinale = duplicateWires
                    //.GroupBy(g => new { g.WireName, g.sheetName, g.WireNumber })
                    .GroupBy(g => new { g.WireName, g.sheetName })
                    .Select(g => new
                    {
                        g.Key.WireName,
                        //g.Key.WireNumber,
                        g.Key.sheetName,
                        Count = g.Count() / 2,  // Assigning count properly
                    }).ToList();

                DuplWireLogging.DeletePreviousLogs();
                foreach (var wire in wireDuplicateFinale)
                {
                    DuplWireLogging.Info($"{wire.WireName} - {wire.sheetName} - {wire.Count} times");
                }
               // duplicatemessage = true;

                MessageBox.Show(
                    $"Duplicate Wires were found in the project.\nPlease find the logs at {GlobalVar.StrtCmd}Logs\\DuplicateWire_{DateTime.Now:yyyy-MM-dd}.log",
                    "Duplicate Wires Found",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation
                );

                Application.Exit();

                //throw new DuplicatWireException();
            }
        }
    }
}
