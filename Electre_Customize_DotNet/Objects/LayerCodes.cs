namespace Electre_Customize_DotNet.Objects
{
    /// <summary>
    /// Electre layer number to NERD letter used on cable/continuity reports.
    /// </summary>
    internal static class LayerCodes
    {
        public static string ToNerd(string layerNumber)
        {
            switch (layerNumber)
            {
                case "158":
                    return "D";
                case "152":
                    return "R";
                case "160":
                case "154":
                    return "N";
                case "169":
                    return "E";
                default:
                    return "";
            }
        }
    }
}
