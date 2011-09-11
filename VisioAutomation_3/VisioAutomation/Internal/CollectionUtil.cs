using System;
using System.Collections.Generic;

namespace VisioAutomation.Internal
{
    internal static class CollectionUtil
    {
        private struct TempSortData<TValue, TKey>
        {
            public TValue Value;
            public TKey Key;

            public TempSortData(TValue v, TKey k)
            {
                this.Value = v;
                this.Key = k;
            }
        }
    }
}
