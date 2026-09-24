using Microsoft.Office.Interop.Excel;

namespace Electre_Customize_DotNet.Helpers.CableList
{
    internal static class CableListHeaderFormulas
    {
        public static void ApplyCl1Formulas(Worksheet ws)
        {
            object[,] cFormulas = new object[5, 1];
            cFormulas[0, 0] = "='CL-1'!C48";
            cFormulas[1, 0] = "='CL-1'!C49";
            cFormulas[2, 0] = "='CL-1'!C50";
            cFormulas[3, 0] = "='CL-1'!C51";
            cFormulas[4, 0] = "='CL-1'!C52";
            ws.Range["C48:C52"].Formula = cFormulas;

            object[,] dFormulas = new object[5, 1];
            dFormulas[0, 0] = "='CL-1'!D48";
            dFormulas[1, 0] = "='CL-1'!D49";
            dFormulas[2, 0] = "='CL-1'!D50";
            dFormulas[3, 0] = "='CL-1'!D51";
            dFormulas[4, 0] = "='CL-1'!D52";
            ws.Range["D48"].NumberFormat = "General";
            ws.Range["D48:D52"].Formula = dFormulas;
            ws.Range["C48:C52"].NumberFormat = "@";
            ws.Range["D48:D52"].NumberFormat = "dd-mm-yyyy";

            ws.Range["F49"].Formula = "='CL-1'!F49";
            ws.Range["F51"].Formula = "='CL-1'!F51";
            ws.Range["F52"].Formula = "='CL-1'!F52";
            ws.Range["K51"].Formula = "='CL-1'!K51";
            ws.Range["K52"].Formula = "='CL-1'!K52";
        }
    }
}
