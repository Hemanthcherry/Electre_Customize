using System.Collections.Generic;
using Electre_Customize_DotNet.Objects;

namespace Electre_Customize_DotNet.Helpers.Global
{
    /// <summary>
    /// Holds one ElectreTraversalIndex for a source list and rebuilds it only when the list REFERENCE
    /// changes. The project-wide collection is reassigned once per data load (not per report), so every
    /// trace call in a report-generation session can share one index instead of rescanning the whole
    /// collection. A new data load produces a new list object, which invalidates the cache automatically.
    /// </summary>
    internal sealed class TraversalIndexCache
    {
        private ElectreTraversalIndex _index;
        private List<ElectreObject> _source;

        public ElectreTraversalIndex For(List<ElectreObject> source)
        {
            if (!ReferenceEquals(_source, source))
            {
                _index = new ElectreTraversalIndex(source);
                _source = source;
            }
            return _index;
        }
    }
}
