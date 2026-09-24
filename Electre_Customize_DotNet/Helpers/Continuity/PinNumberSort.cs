using System.Linq;

namespace Electre_Customize_DotNet.Helpers.Continuity
{
    internal static class PinNumberSort
    {
        /// <summary>Numeric part at the start of a pin number, used as a sort key. Non-numeric pins sort last.</summary>
        public static int ExtractNumericPrefix(string pinNumber)
        {
            string numericPart = new string(pinNumber.TakeWhile(char.IsDigit).ToArray());
            return int.TryParse(numericPart, out int result) ? result : int.MaxValue;
        }
    }
}
