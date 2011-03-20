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

        // Schwartzian transform
        public static IEnumerable<T> GetSortedItems<T, TSortBy>(
            IList<T> list,
            Func<T, TSortBy> get_sortby,
            Comparison<TSortBy> comparison) where TSortBy : IComparable
        {
            var temp_list = new List<TempSortData<T, TSortBy>>(list.Count);
            foreach (var item in list)
            {
                var temp_rec = new TempSortData<T, TSortBy>(item, get_sortby(item));
                temp_list.Add(temp_rec);
            }

            temp_list.Sort((a, b) => comparison(a.Key, b.Key));

            foreach (var item in temp_list)
            {
                yield return item.Value;
            }
        }

        // Schwartzian transform where the key comes from another collection
        public static IEnumerable<T> GetSortedItemsIndexed<T, TSortBy>(
            IList<T> list,
            Func<int, TSortBy> get_sortby,
            Comparison<TSortBy> comparison) where TSortBy : IComparable
        {
            var temp_list = new List<TempSortData<T, TSortBy>>(list.Count);
            for (int i = 0; i < list.Count; i++)
            {
                var temp_rec = new TempSortData<T, TSortBy>(list[i], get_sortby(i));
                temp_list.Add(temp_rec);
            }

            temp_list.Sort((a, b) => comparison(a.Key, b.Key));

            foreach (var item in temp_list)
            {
                yield return item.Value;
            }
        }
    }
}
