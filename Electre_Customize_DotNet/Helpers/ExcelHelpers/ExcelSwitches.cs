using System;
using System.Configuration;

namespace Electre_Customize_DotNet.Helpers.ExcelHelpers
{
    /// <summary>
    /// Opt-in performance switches read from App.config. Every switch is OFF unless set to "true", and OFF
    /// means the original behaviour, step for step.
    /// </summary>
    internal static class ExcelSwitches
    {
        /// <summary>
        /// OnePassWirelistWorkbook: build each sheet / component wire-list workbook in a single pass (create, header, data,
        /// format, Excel sort, ONE SaveAs) instead of SaveAs, re-Open, several Saves and a final Close(true).
        /// </summary>
        public static bool OnePassWirelistWorkbook
        {
            get { return string.Equals(ConfigurationManager.AppSettings["OnePassWirelistWorkbook"], "true", StringComparison.OrdinalIgnoreCase); }
        }
    }
}
