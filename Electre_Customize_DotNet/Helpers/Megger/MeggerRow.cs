#nullable disable
namespace Electre_Customize_DotNet.Helpers.Megger
{
    /// <summary>
    /// One filtered Megger source row. Kept independent of ElectreObject so the Megger engine
    /// (MeggerHighPairs / MeggerFileWriter) has no dependency on the rest of the application.
    /// </summary>
    internal readonly struct MeggerRow
    {
        public readonly string Connector;
        public readonly string Pin;
        public readonly string Wire;
        public readonly string SubNet;

        public MeggerRow(string connector, string pin, string wire, string subNet)
        {
            Connector = connector;
            Pin = pin;
            Wire = wire;
            SubNet = subNet;
        }
    }
}
