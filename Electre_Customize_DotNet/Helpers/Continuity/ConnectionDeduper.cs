namespace Electre_Customize_DotNet.Helpers.Continuity
{
    /// <summary>
    /// Drop reverse-duplicate From/To rows. Continuity copies 7 columns and trims;
    /// Megger copies 4 columns and keeps the original array length.
    /// </summary>
    internal static class ConnectionDeduper
    {
        public static string[,] Remove(string[,] arrFTcwob, int copyCols, bool skipEmptyHashPair, bool trimResult)
        {
            int totalRows = arrFTcwob.GetLength(0);
            string[,] dest = new string[totalRows, copyCols];

            var redundantCheck = new HashSet<string>();
            var duplicateDestTracker = new HashSet<string>();
            var destPinCounts = new Dictionary<string, int>();

            int destCount = 0;

            for (int i = 0; i < totalRows; i++)
            {
                string destKey = $"{arrFTcwob[i, 2]}#{arrFTcwob[i, 3]}";
                if (!destPinCounts.TryAdd(destKey, 1))
                    destPinCounts[destKey]++;
            }

            for (int i = 0; i < totalRows; i++)
            {
                string src = $"{arrFTcwob[i, 0]}#{arrFTcwob[i, 1]}";
                string destKey = $"{arrFTcwob[i, 2]}#{arrFTcwob[i, 3]}";

                string presentLine = $"{src}##{destKey}";
                string connectLine = $"{destKey}##{src}";

                if (destPinCounts[destKey] > 1 && duplicateDestTracker.Contains(destKey))
                    continue;

                if (redundantCheck.Contains(presentLine) || redundantCheck.Contains(connectLine))
                    continue;

                if (skipEmptyHashPair && (presentLine == "####" || connectLine == "####"))
                    continue;

                for (int col = 0; col < copyCols; col++)
                    dest[destCount, col] = arrFTcwob[i, col];

                destCount++;
                redundantCheck.Add(presentLine);

                if (destPinCounts[destKey] > 1)
                    duplicateDestTracker.Add(destKey);
            }

            if (!trimResult)
                return dest;

            string[,] result = new string[destCount, copyCols];
            for (int i = 0; i < destCount; i++)
                for (int j = 0; j < copyCols; j++)
                    result[i, j] = dest[i, j];
            return result;
        }
    }
}
