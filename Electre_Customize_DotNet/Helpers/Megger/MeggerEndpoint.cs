namespace Electre_Customize_DotNet.Helpers.Megger
{
    internal readonly struct MeggerEndpoint
    {
        public readonly string Conn;
        public readonly string Pin;
        public readonly string Wire;
        public readonly string SubNet;
        public readonly bool HasWire;

        public MeggerEndpoint(string conn, string pin, string wire, string subNet, bool hasWire)
        {
            Conn = conn ?? "";
            Pin = pin ?? "";
            Wire = wire;
            SubNet = subNet;
            HasWire = hasWire;
        }
    }
}
