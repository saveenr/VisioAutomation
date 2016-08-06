using System.Collections.Generic;
using System.Linq;

namespace VisioAutomationSamples
{
    internal static class LinqUtil
    {
        public static List<List<T>> Split<T>(List<T> source, int chunksize)
        {
            return source
                .Select((x, i) => new {Index = i, Value = x})
                .GroupBy(x => x.Index/chunksize)
                .Select(x => x.Select(v => v.Value).ToList())
                .ToList();
        }
    }
}