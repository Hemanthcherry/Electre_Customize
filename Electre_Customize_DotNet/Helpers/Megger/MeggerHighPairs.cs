#nullable disable
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Electre_Customize_DotNet.Helpers.Megger
{
    /// <summary>
    /// The "High Megger" pair set, represented as a compact plan instead of materialised lines.
    ///
    /// The original code built one string per pair in a ConcurrentBag, kept a ConcurrentDictionary
    /// key per pair to de-duplicate A-B / B-A, then sorted the whole bag. That is O(pairs) memory,
    /// which is impossible when the output runs to hundreds of millions of lines.
    ///
    /// Here only the distinct (connector, pin) nodes are held (tiny). Lines are generated on demand,
    /// directly as UTF-8 bytes, already in final file order, so any contiguous slice of the output
    /// can be produced independently (which is what lets several files be written in parallel).
    ///
    /// Semantics reproduced from the original MeggerSheet5:
    ///  - a pair of filtered nodes is High unless every row of both nodes shares one (WireNumber, SubNet);
    ///  - such a High pair is dropped if lowMeggerData already contains it;
    ///  - every pair with at least one library-only pin is High and has no low check;
    ///  - each unordered pair is emitted once (A-B / B-A de-duplicated);
    ///  - a filtered node with several rows that disagree on (WireNumber, SubNet) yields one
    ///    "X;pin;X;pin" self line, exactly as the original did for duplicate rows.
    /// Lines come out sorted ascending by (from connector, from pin) using the same culture-aware
    /// comparer the original used - the original sorted descending but re-created a ConcurrentBag from
    /// the sorted sequence, which enumerates in reverse, so the files were in fact ascending.
    /// </summary>
    internal sealed class MeggerHighPairs
    {
        private static readonly byte[] Suffix = Encoding.UTF8.GetBytes(";High Megger" + Environment.NewLine);

        private readonly int _n;
        private readonly byte[][] _prefix;     // UTF-8 of "connector;pin"
        private readonly bool[] _isFiltered;   // node comes from filteredData (as opposed to library only)
        private readonly int[] _uniform;       // id of the single (Wire, SubNet) all its rows share; -1 if they disagree
        private readonly bool[] _selfPair;     // >= 2 rows that disagree -> emits the self line
        private readonly int[][] _lowAdj;      // neighbours (both orientations) present in lowMeggerData
        private readonly long[] _cum;          // _cum[p] = high lines emitted before row p

        public long TotalLines => _cum[_n];
        public int NodeCount => _n;

        private MeggerHighPairs(byte[][] prefix, bool[] isFiltered, int[] uniform, bool[] selfPair, int[][] lowAdj)
        {
            _n = prefix.Length;
            _prefix = prefix;
            _isFiltered = isFiltered;
            _uniform = uniform;
            _selfPair = selfPair;
            _lowAdj = lowAdj;
            _cum = new long[_n + 1];
        }

        public static MeggerHighPairs Build(
            IReadOnlyList<MeggerRow> filtered,
            IReadOnlyList<string> lowLines,
            IReadOnlyList<(string ConnectorName, string PinNumber)> library)
        {
            // ---- 1. distinct nodes, in first-seen order --------------------------------------
            var idOf = new Dictionary<(string, string), int>();
            var conn = new List<string>();
            var pin = new List<string>();
            var isF = new List<bool>();
            var wsOfNode = new List<int>();
            var count = new List<int>();
            var mixed = new List<bool>();
            var wsIds = new Dictionary<(string, string), int>();

            for (int r = 0; r < filtered.Count; r++)
            {
                var row = filtered[r];
                var key = (row.Connector ?? string.Empty, row.Pin ?? string.Empty);
                var ws = (row.Wire, row.SubNet);
                if (!wsIds.TryGetValue(ws, out int wsId))
                {
                    wsId = wsIds.Count;
                    wsIds[ws] = wsId;
                }

                if (!idOf.TryGetValue(key, out int id))
                {
                    id = conn.Count;
                    idOf[key] = id;
                    conn.Add(key.Item1);
                    pin.Add(key.Item2);
                    isF.Add(true);
                    wsOfNode.Add(wsId);
                    count.Add(1);
                    mixed.Add(false);
                }
                else
                {
                    count[id]++;
                    if (wsOfNode[id] != wsId)
                    {
                        mixed[id] = true;
                    }
                }
            }

            if (library != null)
            {
                for (int i = 0; i < library.Count; i++)
                {
                    var key = (library[i].ConnectorName ?? string.Empty, library[i].PinNumber ?? string.Empty);
                    if (!idOf.ContainsKey(key))
                    {
                        idOf[key] = conn.Count;
                        conn.Add(key.Item1);
                        pin.Add(key.Item2);
                        isF.Add(false);
                        wsOfNode.Add(-1);
                        count.Add(0);
                        mixed.Add(false);
                    }
                }
            }

            int n = conn.Count;

            // ---- 2. sort nodes ascending by (connector, pin), first-seen order breaks ties ---
            var order = new int[n];
            for (int i = 0; i < n; i++) order[i] = i;
            var cmp = Comparer<string>.Default;
            Array.Sort(order, (x, y) =>
            {
                int c = cmp.Compare(conn[x], conn[y]);
                if (c != 0) return c;
                c = cmp.Compare(pin[x], pin[y]);
                return c != 0 ? c : x.CompareTo(y);
            });

            var newIdOfOld = new int[n];
            for (int newId = 0; newId < n; newId++) newIdOfOld[order[newId]] = newId;

            var prefix = new byte[n][];
            var isFiltered = new bool[n];
            var uniform = new int[n];
            var selfPair = new bool[n];
            for (int newId = 0; newId < n; newId++)
            {
                int old = order[newId];
                prefix[newId] = Encoding.UTF8.GetBytes(conn[old] + ";" + pin[old]);
                isFiltered[newId] = isF[old];
                uniform[newId] = isF[old] && !mixed[old] ? wsOfNode[old] : -1;
                selfPair[newId] = isF[old] && count[old] >= 2 && mixed[old];
            }

            // ---- 3. lowMeggerData adjacency between filtered nodes ---------------------------
            // The original tested the exact "c1;p1;c2;p2" of the orientation it happened to visit
            // first (parallel race). Both orientations are treated as excluded here, which is one of
            // the outcomes the original could produce for such a pair.
            List<int>[] adj = null;
            if (lowLines != null)
            {
                for (int i = 0; i < lowLines.Count; i++)
                {
                    var parts = lowLines[i].Split(';');
                    if (parts.Length < 5) continue;
                    if (!idOf.TryGetValue((parts[0], parts[1]), out int a)) continue;
                    if (!idOf.TryGetValue((parts[2], parts[3]), out int b)) continue;
                    a = newIdOfOld[a];
                    b = newIdOfOld[b];
                    if (!isFiltered[a] || !isFiltered[b]) continue;
                    adj ??= new List<int>[n];
                    (adj[a] ??= new List<int>()).Add(b);
                    if (a != b) (adj[b] ??= new List<int>()).Add(a);
                }
            }

            var lowAdj = new int[n][];
            if (adj != null)
            {
                for (int i = 0; i < n; i++)
                {
                    if (adj[i] != null) lowAdj[i] = adj[i].ToArray();
                }
            }

            var plan = new MeggerHighPairs(prefix, isFiltered, uniform, selfPair, lowAdj);
            plan.ComputeRowCounts();
            return plan;
        }

        private void ComputeRowCounts()
        {
            var rowCount = new long[_n];
            Parallel.For(0, _n,
                () => new bool[_n],
                (p, state, mark) =>
                {
                    rowCount[p] = CountRow(p, mark);
                    return mark;
                },
                _ => { });

            long running = 0;
            for (int p = 0; p < _n; p++)
            {
                _cum[p] = running;
                running += rowCount[p];
            }
            _cum[_n] = running;
        }

        private long CountRow(int p, bool[] mark)
        {
            var adj = _lowAdj[p];
            if (adj != null) for (int i = 0; i < adj.Length; i++) mark[adj[i]] = true;

            long c = 0;
            if (_selfPair[p] && !mark[p]) c++;

            bool pF = _isFiltered[p];
            int up = _uniform[p];
            if (!pF)
            {
                c += _n - 1 - p; // library-only node: every later node qualifies
            }
            else
            {
                for (int q = p + 1; q < _n; q++)
                {
                    if (!_isFiltered[q] || ((up < 0 || up != _uniform[q]) && !mark[q])) c++;
                }
            }

            if (adj != null) for (int i = 0; i < adj.Length; i++) mark[adj[i]] = false;
            return c;
        }

        /// <summary>Writes high lines [from, to) of the whole high-line sequence into the sink.</summary>
        public void WriteRange(long from, long to, MeggerByteSink sink)
        {
            if (to > TotalLines) to = TotalLines;
            if (to <= from) return;

            // first row whose cumulative end lies beyond 'from'
            int lo = 0, hi = _n - 1;
            while (lo < hi)
            {
                int mid = (lo + hi) >> 1;
                if (_cum[mid + 1] > from) hi = mid; else lo = mid + 1;
            }

            long skip = from - _cum[lo];
            long remaining = to - from;
            var mark = new bool[_n];

            for (int p = lo; p < _n && remaining > 0; p++)
            {
                if (_cum[p + 1] == _cum[p]) continue;

                var adj = _lowAdj[p];
                if (adj != null) for (int i = 0; i < adj.Length; i++) mark[adj[i]] = true;

                var a = _prefix[p];
                if (_selfPair[p] && !mark[p])
                {
                    if (skip > 0) skip--;
                    else { sink.WriteLine(a, a, Suffix); remaining--; }
                }

                bool pF = _isFiltered[p];
                int up = _uniform[p];
                for (int q = p + 1; q < _n && remaining > 0; q++)
                {
                    if (pF && _isFiltered[q] && ((up >= 0 && up == _uniform[q]) || mark[q])) continue;
                    if (skip > 0) { skip--; continue; }
                    sink.WriteLine(a, _prefix[q], Suffix);
                    remaining--;
                }

                if (adj != null) for (int i = 0; i < adj.Length; i++) mark[adj[i]] = false;
            }
        }
    }
}
