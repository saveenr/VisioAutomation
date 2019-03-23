using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VASS = VisioAutomation.ShapeSheet;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionQuery : IEnumerable<SectionQueryColumns>
    {
        private IList<SectionQueryColumns> _list { get; }
        private readonly Dictionary<IVisio.VisSectionIndices, SectionQueryColumns> _map_secindex_to_sec_cols;

        public SectionQuery() : base()
        {
            this._list = new List<SectionQueryColumns>();
            this._map_secindex_to_sec_cols = new Dictionary<IVisio.VisSectionIndices, SectionQueryColumns>();
        }

        private static void RestrictToShapesOnly(SurfaceTarget surface)
        {
            if (surface.Shape == null)
            {
                string msg = "Target must be Shape not Page or Master";
                throw new System.ArgumentException(msg);
            }
        }

        public SectionQueryShapeResults<string> GetFormulas(SurfaceTarget surface)
        {
            RestrictToShapesOnly(surface);

            var shapeidpairs = ShapeIdPairs.Create(surface.Shape);
            var cache = this._create_sectionquerycache(shapeidpairs);

            var srcstream = this._build_src_stream(cache);
            var values = surface.GetFormulasU(srcstream);
            var shape_index = 0;
            var shape_cache_item = cache[shape_index];
            var reader = new VASS.Internal.ArraySegmentReader<string>(values);
            var output_for_shape = this._create_output_for_shape(surface.ID16, shape_cache_item, reader);

            return output_for_shape;
        }

        public SectionQueryShapeResults<TResult> GetResults<TResult>(IVisio.Shape shape)
        {
            var surface = new SurfaceTarget(shape);
            return GetResults<TResult>(surface);
        }

        public SectionQueryShapeResults<TResult> GetResults<TResult>(SurfaceTarget surface)
        {
            RestrictToShapesOnly(surface);

            var shapeidpairs = ShapeIdPairs.Create(surface.Shape);
            var cache = this._create_sectionquerycache(shapeidpairs);

            var srcstream = this._build_src_stream(cache);
            const object[] unitcodes = null;
            var values = surface.GetResults<TResult>(srcstream, unitcodes);
            var shape_index = 0;
            var sectioncache = cache[shape_index];
            var reader = new VASS.Internal.ArraySegmentReader<TResult>(values);
            var output_for_shape = this._create_output_for_shape(surface.ID16, sectioncache, reader);
            return output_for_shape;
        }

        public SectionQueryResults<string> GetFormulas(IVisio.Page page, ShapeIdPairs shapeidpairs)
        {
            var surface = new SurfaceTarget(page);
            return this.GetFormulas(surface, shapeidpairs);
        }


        public SectionQueryResults<TResult> GetResults<TResult>(IVisio.Page page, ShapeIdPairs shapeidpairs)
        {
            var surface = new SurfaceTarget(page);
            return this.GetResults<TResult>(surface, shapeidpairs);
        }

        public SectionQueryResults<TResult> GetResults<TResult>(SurfaceTarget surface, ShapeIdPairs shapeidpairs)
        {
            // Store information about the sections we need to query
            var cache = _create_sectionquerycache(shapeidpairs);

            // Perform the query
            var srcstream = this._build_sidsrc_stream(shapeidpairs, cache);
            const object[] unitcodes = null;
            var values = surface.GetResults<TResult>(srcstream, unitcodes);
            var reader = new VASS.Internal.ArraySegmentReader<TResult>(values);
            var results = this._create_outputs_for_shapes(shapeidpairs, cache, reader);
            return results;
        }
        public SectionQueryResults<string> GetFormulas(SurfaceTarget surface, ShapeIdPairs shapeidpairs)
        {
            // Store information about the sections we need to query
            var cache = _create_sectionquerycache(shapeidpairs);

            // Perform the query
            var srcstream = this._build_sidsrc_stream(shapeidpairs, cache);
            var values = surface.GetFormulasU(srcstream);
            var reader = new VASS.Internal.ArraySegmentReader<string>(values);
            var results = this._create_outputs_for_shapes(shapeidpairs, cache, reader);
            return results;
        }

        private SectionQueryCache _create_sectionquerycache(ShapeIdPairs shapeidpairs)
        {
            // Prepare a cache object
            if (this.Count < 1)
            {
                return new SectionQueryCache(0);
            }

            var _cache = new SectionQueryCache();

            // For each shape, for each section find the number of rows
            foreach (var shapeidpair in shapeidpairs)
            {
               
                // For that shape, fill in the section cache for each section that
                // needs to be queried
                var shapecache = new ShapeCache(this.Count);
                foreach (var sec_cols in this)
                {
                    var shapecacheitem = SectionQuery._create_shapesectioncacheitem(shapeidpair, sec_cols.SectionIndex, sec_cols);
                    shapecache.Add(shapecacheitem);
                }

                // For this shape, add the accumulated info into the cache
                _cache.Add(shapecache);
            }

            // Ensure that we have created a cache for eash shapes
            if (shapeidpairs.Count != _cache.Count)
            {
                string msg = string.Format("mismatch in number of shapes and information collected for shapes");
                throw new Exceptions.InternalAssertionException(msg);
            }

            return _cache;
        }


        private SectionQueryResults<T> _create_outputs_for_shapes<T>(ShapeIdPairs shapeidpairs, SectionQueryCache sectioncache, VASS.Internal.ArraySegmentReader<T> segreader)
        {
            var results = new SectionQueryResults<T>();

            for (int pair_index = 0; pair_index < shapeidpairs.Count; pair_index++)
            {
                var pair = shapeidpairs[pair_index];
                var shapecache = sectioncache[pair_index];
                var shaperesults = this._create_output_for_shape((short)pair.ShapeID, shapecache, segreader);
                results.Add(shaperesults);
            }

            return results;
        }

        private SectionQueryShapeResults<T> _create_output_for_shape<T>(short shapeid, ShapeCache shapecacheitems, VASS.Internal.ArraySegmentReader<T> segreader)
        {
            int original_seg_count = segreader.Count;

            if (shapecacheitems==null)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }


            var results_rows = new List<SectionShapeRows<T>>(shapecacheitems.Count);
            foreach (var shapecacheitem in shapecacheitems)
            {
                var sectionshaperows = new SectionShapeRows<T>(shapecacheitem.RowCount, shapeid, shapecacheitem.SectionColumns.SectionIndex);
                results_rows.Add(sectionshaperows);

                int num_cols = shapecacheitem.SectionColumns.Count;
                foreach (int row_index in shapecacheitem.RowIndexes)
                {
                    var cells = segreader.GetNextSegment(num_cols);
                    var sec_res_row = new Row<T>(shapeid, cells);
                    sectionshaperows.Add(sec_res_row);
                }

            }

            var results = new SectionQueryShapeResults<T>(shapeid, results_rows);

            // the difference in the segment count must match the total number of output cells

            int final_seg_count = segreader.Count;
            int segment_count_delta = final_seg_count - original_seg_count;
            int total_cell_count = shapecacheitems.CountCells();

            if (segment_count_delta  != total_cell_count)
            {
                throw new Exceptions.InternalAssertionException("Unexpected cursor");
            }

            return results;
        }

        private Streams.StreamArray _build_src_stream(SectionQueryCache cache)
        {
            int dummy_shapeid = -1;
            int shapeindex = 0;
            int numcells = cache.CountCells();
            var shapecache = cache[shapeindex];
            var srcs = _sidsrcs_for_shape(dummy_shapeid, shapecache).Select(i => i.Src);
            var stream = VASS.Streams.SrcStreamArrayBuilder.Create(numcells,srcs);

            return stream;
        }

        private VASS.Streams.StreamArray _build_sidsrc_stream(ShapeIdPairs shapeidpairs, SectionQueryCache cache)
        {
            int numcells = cache.CountCells();
            var sidsrcs = _sidsrcs_for_shapes(shapeidpairs, cache);
            var stream = VASS.Streams.SidSrcStreamArrayBuilder.Create(numcells,sidsrcs);
            return stream;
        }

        private static IEnumerable<SidSrc> _sidsrcs_for_shapes(ShapeIdPairs shapeidpairs, SectionQueryCache cache)
        {
            foreach (int shape_ord in Enumerable.Range(0,shapeidpairs.Count))
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

        private static IEnumerable<SidSrc> _sidsrcs_for_shape(int shape_id, ShapeCache shapecache)
        {
            foreach (var shapecacheitem in shapecache)
            {
                foreach (int row_index in shapecacheitem.RowIndexes)
                {
                    var cols = shapecacheitem.SectionColumns;
                    var section_index = shapecacheitem.SectionIndex;
                    foreach (var col in cols)
                    {
                        var sidsrc = new VASS.SidSrc(
                            (short)shape_id,
                            (short)section_index,
                            (short)row_index,
                            col.Src.Cell);
                        yield return sidsrc;
                    }
                }
            }
        }

        public static ShapeCacheItem _create_shapesectioncacheitem(ShapeIdPair shapeidpair, IVisio.VisSectionIndices sec_index, SectionQueryColumns sec_cols)
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
                row_count = shapeidpair.Shape.RowCount[(short)sec_index];
            }

            var shapecacheitem = new ShapeCacheItem((short) shapeidpair.ShapeID, sec_index, sec_cols, row_count);
            return shapecacheitem;
        }

        public IEnumerator<SectionQueryColumns> GetEnumerator()
        {
            return this._list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public SectionQueryColumns this[int index] => this._list[index];

        public SectionQueryColumns Add(IVisio.VisSectionIndices sec_index)
        {
            if (this._map_secindex_to_sec_cols.ContainsKey(sec_index))
            {
                string msg = string.Format("Already contains section index {0} (value={1})", sec_index, (int)sec_index);
                throw new System.ArgumentException(msg);
            }

            var sec_cols = new SectionQueryColumns(sec_index);
            this._list.Add(sec_cols);
            this._map_secindex_to_sec_cols[sec_index] = sec_cols;
            return sec_cols;
        }

        public SectionQueryColumns Add(Src src)
        {
            return this.Add((IVisio.VisSectionIndices)src.Section);
        }

        public int Count => this._list.Count;

    }
}