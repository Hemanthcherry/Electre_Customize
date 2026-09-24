using Microsoft.Office.Interop.Excel;

namespace Electre_Customize_DotNet.Helpers.Continuity
{
    internal static class ContinuityHeaderFormulas
    {
        public static void ApplyCwob1Formulas(Worksheet ws)
        {
            object[,] cFormulas = new object[2, 1];
            cFormulas[0, 0] = "='CWOB-1'!C56";
            cFormulas[1, 0] = "='CWOB-1'!C57";
            ws.Range["C56:C57"].Formula = cFormulas;

            ws.Range["F56"].Formula = "='CWOB-1'!F56";
            ws.Range["I56"].Formula = "='CWOB-1'!I56";
            ws.Range["L56"].Formula = "='CWOB-1'!L56";
            ws.Range["F4"].Formula = "='CWOB-1'!F4";
            ws.Range["I4"].Formula = "='CWOB-1'!I4";
        }
    }
}
