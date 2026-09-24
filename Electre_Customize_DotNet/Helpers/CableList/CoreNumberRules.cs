namespace Electre_Customize_DotNet.Helpers.CableList
{
    /// <summary>Validation of the core number stored in an ElectreObject's Core_Part_Number (Tag3).</summary>
    internal static class CoreNumberRules
    {
        /// <summary>True when the text is an integer core number from 1 to 18.</summary>
        public static bool IsValid(string tag3)
        {
            return int.TryParse(tag3, out int coreNumber) && coreNumber >= 1 && coreNumber <= 18;
        }
    }
}
