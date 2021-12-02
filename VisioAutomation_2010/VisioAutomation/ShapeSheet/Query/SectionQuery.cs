using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionQuery : IEnumerable<ColumnGroup>
    {
        private IList<ColumnGroup> _list_section_query_columns { get; }
        private readonly Dictionary<IVisio.VisSectionIndices, ColumnGroup> _map_secindex_to_sec_cols;

        public SectionQuery() : base()
        {
            this._list_section_query_columns = new List<ColumnGroup>();
            this._map_secindex_to_sec_cols = new Dictionary<IVisio.VisSectionIndices, ColumnGroup>();
        }


        public Data.RowGroup<string> GetFormulas(IVisio.Shape visobjtarget)
        {
            var shapeidpairs = Core.ShapeIDPairs.FromShapes(visobjtarget);
            var cache = this._create_sectionquerycache(shapeidpairs);

            var srcstream = this._build_src_stream(cache);
            var values = visobjtarget.GetFormulasU(srcstream);
            var shape_index = 0;
            var shape_cache_item = cache[shape_index];
            var reader = new VisioAutomation.Internal.ArraySegmentEnumerator<string>(values);
            var output_for_shape = this._create_output_for_shape(visobjtarget.ID16, shape_cache_item, reader);

            return output_for_shape;
        }

        public Data.RowGroup<TResult> GetResults<TResult>(IVisio.Shape shape)
        {
            var shapeidpairs = Core.ShapeIDPairs.FromShapes(shape);
            var cache = this._create_sectionquerycache(shapeidpairs);

            var srcstream = this._build_src_stream(cache);
            const object[] unitcodes = null;
            var values = shape.GetResults<TResult>(srcstream, unitcodes);
            var shape_index = 0;
            var sectioncache = cache[shape_index];
            var reader = new VisioAutomation.Internal.ArraySegmentEnumerator<TResult>(values);
            var output_for_shape = this._create_output_for_shape(shape.ID16, sectioncache, reader);
            return output_for_shape;
        }

        public Data.CellValueGroups<string> GetFormulas(IVisio.Page page, Core.ShapeIDPairs shapeidpairs)
        {
            // Store information about the sections we need to query
            var cache = _create_sectionquerycache(shapeidpairs);

            // Perform the query
            var srcstream = this._build_sidsrc_stream(shapeidpairs, cache);
            var values = page.GetFormulasU(srcstream);
            var reader = new VisioAutomation.Internal.ArraySegmentEnumerator<string>(values);
            var results = this._create_outputs_for_shapes(shapeidpairs, cache, reader);
            return results;
        }


        public Data.CellValueGroups<TResult> GetResults<TResult>(IVisio.Page page, Core.ShapeIDPairs shapeidpairs)
        {
            // Store information about the sections we need to query
            var cache = _create_sectionquerycache(shapeidpairs);

            // Perform the query
            var srcstream = this._build_sidsrc_stream(shapeidpairs, cache);
            const object[] unitcodes = null;
            var values = page.GetResults<TResult>(srcstream, unitcodes);
            var reader = new VisioAutomation.Internal.ArraySegmentEnumerator<TResult>(values);
            var results = this._create_outputs_for_shapes(shapeidpairs, cache, reader);
            return results;
        }


        private VisioAutomation.ShapeSheet.Internal.SectionMetadataCache _create_sectionquerycache(Core.ShapeIDPairs shapeidpairs)
        {
            // Prepare a cache object
            if (this.Count < 1)
            {
                return new VisioAutomation.ShapeSheet.Internal.SectionMetadataCache(0);
            }

            var cache = new VisioAutomation.ShapeSheet.Internal.SectionMetadataCache();

            // For each shape, for each section find the number of rows
            foreach (var shapeidpair in shapeidpairs)
            {
                // For that shape, fill in the section cache for each section that
                // needs to be queried
                var shapecache = new Internal.ShapeCache(this.Count);
                foreach (var sec_cols in this)
                {
                    var shapecacheitem = _create_shapesectioncacheitem(shapeidpair, sec_cols.SectionIndex, sec_cols);
                    shapecache.Add(shapecacheitem);
                }

                // For this shape, add the accumulated info into the cache
                cache.Add(shapecache);
            }

            // Ensure that we have created a cache for eash shapes
            if (shapeidpairs.Count != cache.Count)
            {
                string msg = string.Format("mismatch in number of shapes and information collected for shapes");
                throw new Exceptions.InternalAssertionException(msg);
            }

            return cache;
        }


        private Data.CellValueGroups<T> _create_outputs_for_shapes<T>(Core.ShapeIDPairs shapeidpairs,
            VisioAutomation.ShapeSheet.Internal.SectionMetadataCache sectioncache, VisioAutomation.Internal.ArraySegmentEnumerator<T> segreader)
        {
            var results = new Data.CellValueGroups<T>();

            for (int pair_index = 0; pair_index < shapeidpairs.Count; pair_index++)
            {
                var pair = shapeidpairs[pair_index];
                var shapecache = sectioncache[pair_index];
                var output_for_shape = this._create_output_for_shape((short) pair.ShapeID, shapecache, segreader);
                results.Add(output_for_shape);
            }

            return results;
        }

        private Data.RowGroup<T> _create_output_for_shape<T>(short shapeid, VisioAutomation.ShapeSheet.Internal.ShapeCache shapecacheitems,
            VisioAutomation.Internal.ArraySegmentEnumerator<T> segreader)
        {
            int original_seg_count = segreader.Count;

            if (shapecacheitems == null)
            {
                throw new Exceptions.InternalAssertionException();
            }


            var results_rows = new List<Data.Rows<T>>(shapecacheitems.Count);
            foreach (var shapecacheitem in shapecacheitems)
            {
                var secindex = shapecacheitem.ColumnGroup.SectionIndex;
                var sectionshaperows = new Data.Rows<T>(shapecacheitem.RowCount, shapeid, secindex);
                results_rows.Add(sectionshaperows);

                int num_cols = shapecacheitem.ColumnGroup.Count;
                foreach (int row_index in Enumerable.Range(0, shapecacheitem.RowCount))
                {
                    var cells = segreader.GetNextSegment(num_cols);
                    var sec_res_row = new Data.Row<T>(shapeid, secindex, cells);
                    sectionshaperows.Add(sec_res_row);
                }
            }

            var results = new Data.RowGroup<T>(shapeid, results_rows);

            // the difference in the segment count must match the total number of output cells

            int final_seg_count = segreader.Count;
            int segment_count_delta = final_seg_count - original_seg_count;
            int total_cell_count = shapecacheitems.CountCells();

            if (segment_count_delta != total_cell_count)
            {
                throw new Exceptions.InternalAssertionException("Unexpected cursor");
            }

            return results;
        }

        private Streams.StreamArray _build_src_stream(Internal.SectionMetadataCache cache)
        {
            int dummy_shapeid = -1;
            int shapeindex = 0;
            int numcells = cache.CountCells();
            var shapecache = cache[shapeindex];
            var srcs = _sidsrcs_for_shape(dummy_shapeid, shapecache).Select(i => i.Src);
            var stream = Streams.StreamArray.FromSrc(numcells, srcs);

            return stream;
        }

        private Streams.StreamArray _build_sidsrc_stream(Core.ShapeIDPairs shapeidpairs, VisioAutomation.ShapeSheet.Internal.SectionMetadataCache cache)
        {
            int numcells = cache.CountCells();
            var sidsrcs = _sidsrcs_for_shapes(shapeidpairs, cache);
            var stream = Streams.StreamArray.FromSidSrc(numcells, sidsrcs);
            return stream;
        }

        private static IEnumerable<Core.SidSrc> _sidsrcs_for_shapes(Core.ShapeIDPairs shapeidpairs,
            VisioAutomation.ShapeSheet.Internal.SectionMetadataCache cache)
        {
            foreach (int shape_ord in Enumerable.Range(0, shapeidpairs.Count))
            {
                // For each shape add the cells to query
                var pair = shapeidpairs[shape_ord];
                var shapecache = cache[shape_ord];
                var sidsrcs = _sidsrcs_for_shape(pair.ShapeID, shapecache);
                foreach (var sidsrc in sidsrcs)
                {
                    yield return sidsrc;
                }
            }
        }

        private static IEnumerable<Core.SidSrc> _sidsrcs_for_shape(int shapeid, VisioAutomation.ShapeSheet.Internal.ShapeCache shapecache)
        {
            foreach (var shapecacheitem in shapecache)
            {
                foreach (int row_index in Enumerable.Range(0, shapecacheitem.RowCount))
                {
                    var cols = shapecacheitem.ColumnGroup;
                    var section_index = shapecacheitem.SectionIndex;
                    foreach (var col in cols)
                    {
                        var sidsrc = new Core.SidSrc(
                            (short) shapeid,
                            (short) section_index,
                            (short) row_index,
                            col.Src.Cell);
                        yield return sidsrc;
                    }
                }
            }
        }

        private static Internal.ShapeCacheItem _create_shapesectioncacheitem(Core.ShapeIDPair shapeidpair,
            IVisio.VisSectionIndices sec_index, ColumnGroup sec_cols)
        {
            // first count the rows in the section

            int row_count = 0;
            // For visSectionObject we know the result is always going to be 1
            // so avoid making the call to RowCount[]
            if (sec_index == IVisio.VisSectionIndices.visSectionObject)
            {
                row_count = 1;
            }
            else
            {
                // For all other cases use RowCount[]
                row_count = shapeidpair.Shape.RowCount[(short) sec_index];
            }

            var shapecacheitem = new VisioAutomation.ShapeSheet.Internal.ShapeCacheItem((short) shapeidpair.ShapeID, sec_index, sec_cols, row_count);
            return shapecacheitem;
        }

        public IEnumerator<ColumnGroup> GetEnumerator()
        {
            return this._list_section_query_columns.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public ColumnGroup this[int index] => this._list_section_query_columns[index];

        public ColumnGroup Add(IVisio.VisSectionIndices sec_index)
        {
            if (this._map_secindex_to_sec_cols.ContainsKey(sec_index))
            {
                string msg = string.Format("Already contains section index {0} (value={1})", sec_index,
                    (int) sec_index);
                throw new System.ArgumentException(msg);
            }

            var sec_cols = new ColumnGroup(sec_index);
            this._list_section_query_columns.Add(sec_cols);
            this._map_secindex_to_sec_cols[sec_index] = sec_cols;
            return sec_cols;
        }

        public ColumnGroup Add(Core.Src src)
        {
            return this.Add((IVisio.VisSectionIndices) src.Section);
        }

        public int Count => this._list_section_query_columns.Count;
    }
}