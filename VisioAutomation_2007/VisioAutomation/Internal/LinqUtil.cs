using System;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Internal
{
    internal static class LinqUtil
    {

        public enum ChunkCommand
        {
            Start,
            Add,
            End
        }

        public struct ChunkedList<T,TCat>
        {
            public readonly TCat Category;
            public readonly List<T> Items;

            public ChunkedList(List<T> items, TCat cat)
            {
                Items = items;
                Category = cat;
            }
        }

        public struct ChunkRecord<T,TCat>
        {
            private T _item;
            private TCat _cat;
            private ChunkCommand _Command;

            private ChunkRecord(ChunkCommand cmd, T item, TCat cat)
            {
                _item = item;
                _Command = cmd;
                _cat = cat;
            }

            public TCat Category { get { return this._cat;}}

            public T Item
            {
                get
                {
                    if (this.Command == ChunkCommand.Add)
                    {
                        return this._item;
                    }
                    else
                    {
                        throw new Exception();
                    }
                }
            }
            
            public ChunkCommand Command { get { return this._Command; }}

            internal static ChunkRecord<T,TCat> GetStartRecord(TCat cat)
            {
                return new ChunkRecord<T, TCat>(ChunkCommand.Start, default(T), cat);                
            }

            internal static ChunkRecord<T, TCat> GetEndRecord(TCat cat)
            {
                return new ChunkRecord<T, TCat>(ChunkCommand.End, default(T), cat);
            }

            internal static ChunkRecord<T, TCat> GetAddRecord(TCat cat, T item)
            {
                return new ChunkRecord<T, TCat>(ChunkCommand.Add, item, cat);
            }
        }

        public static IEnumerable<ChunkRecord<T,TCat>> ChunkByCategory<T,TCat>(IEnumerable<T> items, System.Func<T, TCat> func_categorize, System.Func<TCat,TCat,bool> fund_cat_eq)
        {
            TCat old_cat = default(TCat);
            int state = 0;
            foreach (var item in items)
            {
                TCat item_category = func_categorize(item);
                if (state == 0)
                {
                    // state 0 means we've never encountered any record before so we must START and ADD
                    yield return ChunkRecord<T, TCat>.GetStartRecord(item_category);
                    yield return ChunkRecord<T, TCat>.GetAddRecord(item_category,item);

                    old_cat = item_category;
                    state = 1;
                }
                else if (state ==1)
                {
                    // we are in the middle of a sequence
                    if (fund_cat_eq(item_category,old_cat))
                    {
                        // they have the same category so continue to ADD
                        yield return ChunkRecord<T, TCat>.GetAddRecord(item_category, item);
                    }
                    else
                    {
                        // they have the same category so continue so END, then START, then ADD
                        yield return ChunkRecord<T, TCat>.GetEndRecord(old_cat);

                        yield return ChunkRecord<T, TCat>.GetStartRecord(item_category);
                        yield return ChunkRecord<T, TCat>.GetAddRecord(item_category, item);

                        old_cat = item_category;
                    }            
                }
            }
            if (state == 0)
            {
                // we never encountered any items so, do nothing
            }
            else
            {
                // there is a sequence still open, END it
                yield return ChunkRecord<T, TCat>.GetEndRecord(old_cat);
            }
        }


        public static IEnumerable<ChunkedList<T,bool>> ChunkByBool<T>( IEnumerable<T> items, System.Func<T,bool> func_categorize)
        {
            var true_col = new List<T>();
            var false_col = new List<T>();

            foreach (var cmd in ChunkByCategory(items, i => func_categorize(i), (a, b) => a == b))
            {
                if (cmd.Command == ChunkCommand.Start)
                {
                    if (cmd.Category == true)
                    {
                        true_col.Clear();
                    }
                    else if (cmd.Category == false)
                    {
                        false_col.Clear();
                    }
                }
                else if (cmd.Command == ChunkCommand.Add)
                {
                    if (cmd.Category == true)
                    {
                        true_col.Add(cmd.Item);
                    }
                    else if (cmd.Category == false)
                    {
                        false_col.Add(cmd.Item);
                    }
                }
                else if (cmd.Command == ChunkCommand.End)
                {
                    if (cmd.Category == true)
                    {
                        yield return new ChunkedList<T,bool>(true_col,true);
                        true_col.Clear();
                    }
                    else if (cmd.Category == false)
                    {
                        yield return new ChunkedList<T,bool>(false_col,false);
                        false_col.Clear();
                    }
                }
            }
        }

    }
}