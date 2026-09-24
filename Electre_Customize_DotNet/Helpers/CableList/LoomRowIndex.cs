using System.Collections.Generic;

namespace Electre_Customize_DotNet.Helpers.CableList
{
    /// <summary>
    /// Index over arrFT_CwithBCProject (the project-wide from/to wire array) that answers
    /// "which rows belong to loom X" without rescanning every row once per selected loom.
    ///
    /// A row belongs to a loom through its FROM bundle (column 8) or its TO bundle (column 10). The
    /// original per-loom loop tested FROM first and only then TO (if / else-if), so:
    ///  - a row is recorded under its FROM bundle as "IsFrom = true";
    ///  - and under its TO bundle as "IsFrom = false", unless that is the same bundle (then it is
    ///    recorded once, as FROM, exactly like the short-circuiting else-if);
    ///  - rows are recorded in ascending row order, so per-loom iteration order is unchanged;
    ///  - if a sheet filter is active, rows on other sheets are skipped (column 7);
    ///  - a null bundle never matches any loom (Dictionary keys cannot be null, and the unfilled
    ///    trailing rows of the array are all null).
    /// </summary>
    internal sealed class LoomRowIndex
    {
        private const int SheetColumn = 7;
        private const int FromBundleColumn = 8;
        private const int ToBundleColumn = 10;

        private readonly Dictionary<string, List<(int Y, bool IsFrom)>> _rowsByLoom = new();

        private LoomRowIndex() { }

        public static LoomRowIndex Build(object[,] projectRows, IEnumerable<string> selectedSheets)
        {
            var index = new LoomRowIndex();
            HashSet<string> sheetFilter = selectedSheets != null ? new HashSet<string>(selectedSheets) : null;

            int totalRows = projectRows.GetLength(0);
            for (int y = 0; y < totalRows; y++)
            {
                string sheetName = projectRows[y, SheetColumn]?.ToString();
                if (sheetFilter != null && !sheetFilter.Contains(sheetName))
                {
                    continue;
                }

                string fromBundle = projectRows[y, FromBundleColumn]?.ToString();
                string toBundle = projectRows[y, ToBundleColumn]?.ToString();

                if (fromBundle != null)
                {
                    index.Add(fromBundle, y, isFrom: true);
                }

                if (toBundle != null && toBundle != fromBundle)
                {
                    index.Add(toBundle, y, isFrom: false);
                }
            }

            return index;
        }

        private void Add(string loomName, int row, bool isFrom)
        {
            if (!_rowsByLoom.TryGetValue(loomName, out var list))
            {
                list = new List<(int, bool)>();
                _rowsByLoom[loomName] = list;
            }
            list.Add((row, isFrom));
        }

        /// <summary>The matching rows of a loom, in ascending row order; false if it has none.</summary>
        public bool TryGetRows(string loomName, out List<(int Y, bool IsFrom)> rows)
        {
            if (loomName == null)
            {
                rows = null;
                return false;
            }
            return _rowsByLoom.TryGetValue(loomName, out rows);
        }
    }
}
