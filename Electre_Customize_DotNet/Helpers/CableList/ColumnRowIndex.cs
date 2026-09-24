using System;
using System.Collections.Generic;

namespace Electre_Customize_DotNet.Helpers.CableList
{
    /// <summary>
    /// "Which rows of this array have value X in column C" without rescanning the whole array once
    /// per selected sheet / component. Equivalent to the original
    /// <c>array[y, C] != null &amp;&amp; array[y, C].ToString() == value</c> scan: ordinal, case-sensitive,
    /// rows returned in ascending order, null cells never match.
    /// </summary>
    internal sealed class ColumnRowIndex
    {
        private static readonly IReadOnlyList<int> NoRows = Array.Empty<int>();

        private readonly Dictionary<string, List<int>> _rowsByValue = new Dictionary<string, List<int>>(StringComparer.Ordinal);

        public ColumnRowIndex(object[,] source, int column)
        {
            int totalRows = source.GetLength(0);
            for (int y = 0; y < totalRows; y++)
            {
                object cell = source[y, column];
                if (cell == null)
                {
                    continue;
                }

                string key = cell.ToString();
                if (key == null)
                {
                    continue;
                }

                if (!_rowsByValue.TryGetValue(key, out var rows))
                {
                    rows = new List<int>();
                    _rowsByValue[key] = rows;
                }
                rows.Add(y);
            }
        }

        public IReadOnlyList<int> RowsFor(string value)
        {
            if (value != null && _rowsByValue.TryGetValue(value, out var rows))
            {
                return rows;
            }
            return NoRows;
        }
    }
}
