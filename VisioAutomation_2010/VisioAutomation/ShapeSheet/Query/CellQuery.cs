using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VASS = VisioAutomation.ShapeSheet;

namespace VisioAutomation.ShapeSheet.Query
{
    public class CellQuery
    {
        public ColumnList Columns { get; }

        public CellQuery()
        {
            this.Columns = new ColumnList(0);
        }

        private static void RestrictToShapesOnly(SurfaceTarget surface)
        {
            if (surface.Shape == null)
            {
                string msg = "Target must be Shape not Page or Master";
                throw new System.ArgumentException(msg);
            }
        }

        public CellOutput<string> GetFormulas(IVisio.Shape shape)
        {
            var surface = new SurfaceTarget(shape);
            return GetFormulas(surface);
        }

        public CellOutput<string> GetCells(IVisio.Shape shape, CellValueType type)
        {
            var surface = new SurfaceTarget(shape);
            if (type == CellValueType.Formula)
            {
                return GetFormulas(surface);
            }
            else
            {
                return GetResults<string>(surface);
            }
        }

        public CellOutput<string> GetFormulas(SurfaceTarget surface)
        {
            RestrictToShapesOnly(surface);

            var srcstream = this._build_src_stream();
            var values = surface.GetFormulasU(srcstream);
            var seg_builder = new Internal.ArraySegmentReader<string>(values);
            var output_for_shape = this._create_output_for_shape(surface.ID16, seg_builder);

            return output_for_shape;
        }

        public CellOutput<TResult> GetResults<TResult>(IVisio.Shape shape)
        {
            var surface = new SurfaceTarget(shape);
            return GetResults<TResult>(surface);
        }

        public CellOutput<TResult> GetResults<TResult>(SurfaceTarget surface)
        {
            RestrictToShapesOnly(surface);

            var srcstream = this._build_src_stream();
            const object[] unitcodes = null;
            var values = surface.GetResults<TResult>(srcstream, unitcodes);
            var seg_builder = new Internal.ArraySegmentReader<TResult>(values);
            var output_for_shape = this._create_output_for_shape(surface.ID16, seg_builder);
            return output_for_shape;
        }

        public CellOutputList<string> GetFormulas(IVisio.Page page, IList<int> shapeids)
        {
            var surface = new SurfaceTarget(page);
            return this.GetFormulas(surface, shapeids);
        }

        public CellOutputList<string> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var surface = new SurfaceTarget(page);
            if (type == CellValueType.Formula)
            {
                return this.GetFormulas(surface, shapeids);
            }
            else
            {
                return this.GetResults<string>(surface, shapeids);
            }
        }


        public CellOutputList<string> GetFormulas(SurfaceTarget surface, IList<int> shapeids)
        {
            var shapes = new List<IVisio.Shape>(shapeids.Count);
            shapes.AddRange(shapeids.Select(shapeid => surface.Shapes.ItemFromID16[(short)shapeid]));

            var srcstream = this._build_sidsrc_stream(shapeids);
            var values = surface.GetFormulasU(srcstream);
            var seg_builder = new Internal.ArraySegmentReader<string>(values);
            var list = this._create_outputs_for_shapes(shapeids, seg_builder);
            return list;
        }

        public CellOutputList<TResult> GetResults<TResult>(IVisio.Page page, IList<int> shapeids)
        {
            var surface = new SurfaceTarget(page);
            return this.GetResults<TResult>(surface, shapeids);
        }

        public CellOutputList<TResult> GetResults<TResult>(SurfaceTarget surface, IList<int> shapeids)
        {
            var shapes = new List<IVisio.Shape>(shapeids.Count);
            shapes.AddRange(shapeids.Select(shapeid => surface.Shapes.ItemFromID16[(short)shapeid]));

            var srcstream = this._build_sidsrc_stream(shapeids);
            const object[] unitcodes = null;
            var values = surface.GetResults<TResult>(srcstream, unitcodes);
            var seg_builder = new Internal.ArraySegmentReader<TResult>(values);
            var list = this._create_outputs_for_shapes(shapeids, seg_builder);
            return list;
        }

        private CellOutputList<T> _create_outputs_for_shapes<T>(IList<int> shapeids, VASS.Internal.ArraySegmentReader<T> segReader)
        {
            var output_for_all_shapes = new CellOutputList<T>();

            for (int shape_index = 0; shape_index < shapeids.Count; shape_index++)
            {
                var shapeid = shapeids[shape_index];
                var output_for_shape = this._create_output_for_shape((short)shapeid, segReader);
                output_for_all_shapes.Add(output_for_shape);
            }

            return output_for_all_shapes;
        }

        private CellOutput<T> _create_output_for_shape<T>(short shapeid, VASS.Internal.ArraySegmentReader<T> segReader)
        {
            int original_seg_size = segReader.Count;

            var output = new CellOutput<T>(shapeid, this.Columns.Count, segReader.GetNextSegment(this.Columns.Count));

            int final_seg_size = segReader.Count;

            if ((final_seg_size - original_seg_size) != output.__totalcellcount)
            {
                throw new Exceptions.InternalAssertionException("Unexpected cursor");
            }

            return output;
        }

        private int _get_total_cell_count(int numshapes)
        {
            return this.Columns.Count * numshapes;
        }

        private Streams.StreamArray _build_src_stream()
        {
            int dummy_shapeid = -1;
            int numshapes = 1;
            int numcells = this._get_total_cell_count(numshapes);
            var stream = new VASS.Streams.SrcStreamArrayBuilder(numcells);
            var cellinfos = this._enum_total_cellinfo(dummy_shapeid);
            var srcs = cellinfos.Select(i => i.SidSrc.Src);
            stream.AddRange(srcs);

            return stream.ToStreamArray();
        }

        private VASS.Streams.StreamArray _build_sidsrc_stream(IList<int> shapeids)
        {
            int numshapes = shapeids.Count;
            int numcells = this._get_total_cell_count(numshapes);

            var stream = new VASS.Streams.SidSrcStreamArrayBuilder(numcells);

            for (int shapeindex = 0; shapeindex < shapeids.Count; shapeindex++)
            {
                // For each shape add the cells to query
                var shapeid = shapeids[shapeindex];

                var cellinfos = this._enum_total_cellinfo(shapeid);
                var sidsrcs = cellinfos.Select(i => i.SidSrc);
                stream.AddRange(sidsrcs);
            }

            return stream.ToStreamArray();
        }

        private IEnumerable<Internal.QueryCellInfo> _enum_total_cellinfo(int shapeid)
        {
            foreach (var col in this.Columns)
            {
                var sidsrc = new SidSrc((short)shapeid, col.Src);
                var cellinfo = new Internal.QueryCellInfo(sidsrc, col);
                yield return cellinfo;
            }
        }
    }
}