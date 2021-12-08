using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Core;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using VisioAutomation.Internal;
using VisioAutomation.ShapeSheet.Data;
using VisioAutomation.ShapeSheet.Internal;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionQuery : IEnumerable<Data.DataColumns>
    {
        private IList<Data.DataColumns> _list_section_query_columns { get; }
        private readonly Dictionary<IVisio.VisSectionIndices, Data.DataColumns> _map_secindex_to_sec_cols;

        public SectionQuery() : base()
        {
            this._list_section_query_columns = new List<Data.DataColumns>();
            this._map_secindex_to_sec_cols = new Dictionary<IVisio.VisSectionIndices, Data.DataColumns>();
        }


        public Data.DataRowGroup<string> GetFormulas(IVisio.Shape visobjtarget)
        {
            var shapeidpairs = Core.ShapeIDPairs.FromShapes(visobjtarget);
            var cache = this._create_section_metadata_cache(shapeidpairs);

            var srcstream = this._build_src_stream(cache);
            var values = visobjtarget.GetFormulasU(srcstream);
            var shape_index = 0;
            var shape_cache_item = cache[shape_index];
            var segments = new VisioAutomation.Internal.ArraySegmentEnumerator<string>(values);
            var output_for_shape = this.segments_to_rowgroup(shape_cache_item, visobjtarget.ID16, segments);

            return output_for_shape;
        }

        public Data.DataRowGroup<TResult> GetResults<TResult>(IVisio.Shape shape)
        {
            var shapeidpairs = Core.ShapeIDPairs.FromShapes(shape);
            var sectionmetadatacache = this._create_section_metadata_cache(shapeidpairs);

            var srcstream = this._build_src_stream(sectionmetadatacache);
            const object[] unitcodes = null;
            var values = shape.GetResults<TResult>(srcstream, unitcodes);
            var shape_index = 0;
            var shapemetadatacache = sectionmetadatacache[shape_index];
            var segments = new VisioAutomation.Internal.ArraySegmentEnumerator<TResult>(values);
            var output_for_shape = this.segments_to_rowgroup(shapemetadatacache, shape.ID16, segments);
            return output_for_shape;
        }

        public Data.DataRowGroups<string> GetFormulas(IVisio.Page page, Core.ShapeIDPairs shapeidpairs)
        {
            // Store information about the sections we need to query
            var cache = _create_section_metadata_cache(shapeidpairs);

            // Perform the query
            var srcstream = this._build_sidsrc_stream(shapeidpairs, cache);
            var values = page.GetFormulasU(srcstream);
            var segments = new VisioAutomation.Internal.ArraySegmentEnumerator<string>(values);
            var results = this._segments_to_rowgroups(shapeidpairs, cache, segments);
            return results;
        }


        public Data.DataRowGroups<TResult> GetResults<TResult>(IVisio.Page page, Core.ShapeIDPairs shapeidpairs)
        {
            // Store information about the sections we need to query
            var cache = _create_section_metadata_cache(shapeidpairs);

            // Perform the query
            var srcstream = this._build_sidsrc_stream(shapeidpairs, cache);
            const object[] unitcodes = null;
            var values = page.GetResults<TResult>(srcstream, unitcodes);
            var segments = new VisioAutomation.Internal.ArraySegmentEnumerator<TResult>(values);
            var results = this._segments_to_rowgroups(shapeidpairs, cache, segments);
            return results;
        }


        private Internal.SectionMetadataCache _create_section_metadata_cache(Core.ShapeIDPairs shapeidpairs)
        {
            // Prepare a cache object
            if (this.Count < 1)
            {
                return new Internal.SectionMetadataCache(0);
            }

            var cache = new Internal.SectionMetadataCache();

            // For each shape, for each section find the number of rows
            foreach (var shapeidpair in shapeidpairs)
            {
                // For that shape, fill in the section cache for each section that
                // needs to be queried
                var shapecache = new Internal.ShapeMetadataCache(this.Count);
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


        private Data.DataRowGroups<T> _segments_to_rowgroups<T>(Core.ShapeIDPairs shapeidpairs,
            Internal.SectionMetadataCache sectionmetadatacache, VisioAutomation.Internal.ArraySegmentEnumerator<T> segreader)
        {
            var datarowgroups = new Data.DataRowGroups<T>();

            for (int pair_index = 0; pair_index < shapeidpairs.Count; pair_index++)
            {
                var pair = shapeidpairs[pair_index];
                var sectionmetadata = sectionmetadatacache[pair_index];
                var output_for_shape = this.segments_to_rowgroup(sectionmetadata, (short) pair.ShapeID, segreader);
                datarowgroups.Add(output_for_shape);
            }

            return datarowgroups;
        }

        private DataRowGroup<T> segments_to_rowgroup<T>(ShapeMetadataCache shapecacheitems,
            short shapeid,
            ArraySegmentEnumerator<T> segreader)
        {
            int original_seg_count = segreader.Count;

            if (shapecacheitems == null)
            {
                throw new Exceptions.InternalAssertionException();
            }


            var results_rows = new List<Data.DataRows<T>>(shapecacheitems.Count);
            foreach (var shapecacheitem in shapecacheitems)
            {
                var secindex = shapecacheitem.ColumnGroup.SectionIndex;
                var datarows = new Data.DataRows<T>(shapecacheitem.RowCount, shapeid, secindex);
                results_rows.Add(datarows);

                int num_cols = shapecacheitem.ColumnGroup.Count();
                foreach (int row_index in Enumerable.Range(0, shapecacheitem.RowCount))
                {
                    var cells = segreader.GetNextSegment(num_cols);
                    var sec_res_row = new Data.DataRow<T>(shapeid, secindex, cells);
                    datarows.Add(sec_res_row);
                }
            }

            var results = new Data.DataRowGroup<T>(shapeid, results_rows);

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
            var srcs = _sidsrcs_for_shape(shapecache, dummy_shapeid).Select(i => i.Src);
            var stream = Streams.StreamArray.FromSrc(numcells, srcs);

            return stream;
        }

        private Streams.StreamArray _build_sidsrc_stream(Core.ShapeIDPairs shapeidpairs, Internal.SectionMetadataCache cache)
        {
            int numcells = cache.CountCells();
            var sidsrcs = _sidsrcs_for_shapes(shapeidpairs, cache);
            var stream = Streams.StreamArray.FromSidSrc(numcells, sidsrcs);
            return stream;
        }

        private static IEnumerable<Core.SidSrc> _sidsrcs_for_shapes(Core.ShapeIDPairs shapeidpairs,
            Internal.SectionMetadataCache sectionmetadatacache)
        {
            foreach (int shape_ord in Enumerable.Range(0, shapeidpairs.Count))
            {
                // For each shape add the cells to query
                var pair = shapeidpairs[shape_ord];
                var shapemetadatacache = sectionmetadatacache[shape_ord];
                var sidsrcs = _sidsrcs_for_shape(shapemetadatacache, pair.ShapeID);
                foreach (var sidsrc in sidsrcs)
                {
                    yield return sidsrc;
                }
            }
        }

        private static IEnumerable<SidSrc> _sidsrcs_for_shape(ShapeMetadataCache shapemetadatacache, int shapeid)
        {
            foreach (var shapecacheitem in shapemetadatacache)
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

        private static Internal.ShapeMetadataCacheItem _create_shapesectioncacheitem(Core.ShapeIDPair shapeidpair,
            IVisio.VisSectionIndices sec_index, Data.DataColumns sec_cols)
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

            var shapecacheitem = new Internal.ShapeMetadataCacheItem((short) shapeidpair.ShapeID, sec_index, sec_cols, row_count);
            return shapecacheitem;
        }

        public IEnumerator<Data.DataColumns> GetEnumerator()
        {
            return this._list_section_query_columns.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public Data.DataColumns this[int index] => this._list_section_query_columns[index];

        public Data.DataColumns Add(IVisio.VisSectionIndices sec_index)
        {
            if (this._map_secindex_to_sec_cols.ContainsKey(sec_index))
            {
                string msg = string.Format("Already contains section index {0} (value={1})", sec_index,
                    (int) sec_index);
                throw new System.ArgumentException(msg);
            }

            var sec_cols = new Data.DataColumns(sec_index);
            this._list_section_query_columns.Add(sec_cols);
            this._map_secindex_to_sec_cols[sec_index] = sec_cols;
            return sec_cols;
        }

        public Data.DataColumns Add(Core.Src src)
        {
            return this.Add((IVisio.VisSectionIndices) src.Section);
        }

        public int Count => this._list_section_query_columns.Count;
    }
}