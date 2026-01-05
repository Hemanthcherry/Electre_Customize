using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Electre_Customize_DotNet.Objects
{
    public class ElectreObject
    {
        // Public properties with auto-implemented getters and setters
        public string SheetName { get; set; }
        public string SheetNumber { get; set; }
        public string DrawingNumber { get; set; }
        public string DefaultGauge { get; set; }
        public string BundleName { get; set; }
        public string EquipmentName { get; set; }
        public string ConnectorName { get; set; }
        public string PinNumber { get; set; }
        public string Ends { get; set; }
        public string FunctionalDesignation { get; set; }
        public string SymbolName { get; set; }
        public string ComponentType { get; set; }
        public string WireNumber { get; set; }
        public string Signal { get; set; }
        public string Group { get; set; }
        public string Gauge { get; set; }
        public string CableType { get; set; }
        public string Length { get; set; }
        public string Layer { get; set; }
        public string OverShield { get; set; }
        public string Net { get; set; }
        public string SubNet { get; set; }
        public string Shunt { get; set; }
        public string Tag1 { get; set; }
        public string Tag2 { get; set; }
        public string Core_Part_Number { get; set; }
        public string Tag4 { get; set; }
        public string Voltage { get; set; }
        public string Tag6_Link { get; set; }
        public string Tag7 { get; set; }
        public string ShuntExt1 { get; set; }
        public string ShuntExt1Address { get; set; }
        public string Panel { get; set; }
        public string NoMegger { get; set; }
        //public string Tag8 { get; set; }    

        // Commented out properties
        /*
        public string ToConnectorSHEET { get; set; }
        public string ToConnectorLOOM { get; set; }
        */
    }
}
