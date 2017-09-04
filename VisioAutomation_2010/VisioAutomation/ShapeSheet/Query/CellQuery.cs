using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Query
{
    public class CellQuery
    {
        public CellColumnCollection Cells { get; }

        public CellQuery()
        {
            this.Cells = new CellColumnCollection(0);
        }

        public CellColumn AddCell(ShapeSheet.Src src, string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            var col = this.Cells.Add(src, name);
            return col;
        }

        private static void RestrictToShapesOnly(SurfaceTarget surface)
        {
            if (surface.Shape == null)
            {
                string msg = "Target must be Shape not Page or Master";
                throw new System.ArgumentException(msg);
            }
        }

        public QueryOutputCells<string> GetFormulas(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var surface = new SurfaceTarget(shape);
            return GetFormulas(surface);
        }

        public QueryOutputCells<string> GetFormulas(SurfaceTarget surface)
        {
            RestrictToShapesOnly(surface);

            var srcstream = this._build_src_stream();
            var values = surface.GetFormulasU(srcstream);
            var shape_index = 0;
            var seg_builder = new VisioAutomation.Utilities.ArraySegmentReader<string>(values);
            var output_for_shape = this._create_output_for_shape(surface.ID16, seg_builder);

            return output_for_shape;
        }

        public QueryOutputCells<TResult> GetResults<TResult>(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var surface = new SurfaceTarget(shape);
            return GetResults<TResult>(surface);
        }

        public QueryOutputCells<TResult> GetResults<TResult>(SurfaceTarget surface)
        {
            RestrictToShapesOnly(surface);

            var srcstream = this._build_src_stream();
            const object[] unitcodes = null;
            var values = surface.GetResults<TResult>(srcstream, unitcodes);
            var shape_index = 0;
            var seg_builder = new VisioAutomation.Utilities.ArraySegmentReader<TResult>(values);
            var output_for_shape = this._create_output_for_shape(surface.ID16, seg_builder);
            return output_for_shape;
        }

        public QueryOutputCells<ShapeSheet.CellData> GetFormulasAndResults(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var surface = new SurfaceTarget(shape);
            return this.GetFormulasAndResults(surface);
        }

        public QueryOutputCells<ShapeSheet.CellData> GetFormulasAndResults(SurfaceTarget surface)
        {
            RestrictToShapesOnly(surface);

            var srcstream = this._build_src_stream();
            const object[] unitcodes = null;
            var formulas = surface.GetFormulasU(srcstream);
            var results = surface.GetResults<string>(srcstream, unitcodes);
            var combined_data = QueryUtil._combine_formulas_and_results(formulas, results);

            var seg_builder = new VisioAutomation.Utilities.ArraySegmentReader<CellData>(combined_data);
            var output_for_shape = this._create_output_for_shape(surface.ID16, seg_builder);
            return output_for_shape;
        }

        public QueryOutputCollectionCells<string> GetFormulas(Microsoft.Office.Interop.Visio.Page page, IList<int> shapeids)
        {
            var surface = new SurfaceTarget(page);
            return this.GetFormulas(surface, shapeids);
        }

        public QueryOutputCollectionCells<string> GetFormulas(SurfaceTarget surface, IList<int> shapeids)
        {
            var shapes = new List<Microsoft.Office.Interop.Visio.Shape>(shapeids.Count);
            shapes.AddRange(shapeids.Select(shapeid => surface.Shapes.ItemFromID16[(short)shapeid]));

            var srcstream = this._build_sidsrc_stream(shapeids);
            var values = surface.GetFormulasU(srcstream);
            var seg_builder = new VisioAutomation.Utilities.ArraySegmentReader<string>(values);
            var list = this._create_outputs_for_shapes(shapeids, null, seg_builder);
            return list;
        }

        public QueryOutputCollectionCells<TResult> GetResults<TResult>(Microsoft.Office.Interop.Visio.Page page, IList<int> shapeids)
        {
            var surface = new SurfaceTarget(page);
            return this.GetResults<TResult>(surface, shapeids);
        }

        public QueryOutputCollectionCells<TResult> GetResults<TResult>(SurfaceTarget surface, IList<int> shapeids)
        {
            var shapes = new List<Microsoft.Office.Interop.Visio.Shape>(shapeids.Count);
            shapes.AddRange(shapeids.Select(shapeid => surface.Shapes.ItemFromID16[(short)shapeid]));

            var srcstream = this._build_sidsrc_stream(shapeids);
            const object[] unitcodes = null;
            var values = surface.GetResults<TResult>(srcstream, unitcodes);
            var seg_builder = new VisioAutomation.Utilities.ArraySegmentReader<TResult>(values);
            var list = this._create_outputs_for_shapes(shapeids, null, seg_builder);
            return list;
        }

        public QueryOutputCollectionCells<ShapeSheet.CellData> GetFormulasAndResults(Microsoft.Office.Interop.Visio.Page page, IList<int> shapeids)
        {
            var surface = new SurfaceTarget(page);
            return this.GetFormulasAndResults(surface, shapeids);
        }

        public QueryOutputCollectionCells<ShapeSheet.CellData> GetFormulasAndResults(SurfaceTarget surface, IList<int> shapeids)
        {
            var shapes = new List<Microsoft.Office.Interop.Visio.Shape>(shapeids.Count);
            shapes.AddRange(shapeids.Select(shapeid => surface.Shapes.ItemFromID16[(short)shapeid]));

            var srcstream = this._build_sidsrc_stream(shapeids);
            const object[] unitcodes = null;
            var results = surface.GetResults<string>(srcstream, unitcodes);
            var formulas = surface.GetFormulasU(srcstream);
            var combined_data = QueryUtil._combine_formulas_and_results(formulas, results);

            var seg_builder = new VisioAutomation.Utilities.ArraySegmentReader<CellData>(combined_data);
            var r = this._create_outputs_for_shapes(shapeids, null, seg_builder);
            return r;
        }

        private QueryOutputCollectionCells<T> _create_outputs_for_shapes<T>(IList<int> shapeids, SectionInfoCache cache, VisioAutomation.Utilities.ArraySegmentReader<T> segReader)
        {
            var output_for_all_shapes = new QueryOutputCollectionCells<T>();

            for (int shape_index = 0; shape_index < shapeids.Count; shape_index++)
            {
                var shapeid = shapeids[shape_index];
                var output_for_shape = this._create_output_for_shape((short)shapeid, segReader);
                output_for_all_shapes.Add(output_for_shape);
            }

            return output_for_all_shapes;
        }

        private QueryOutputCells<T> _create_output_for_shape<T>(short shapeid, VisioAutomation.Utilities.ArraySegmentReader<T> segReader)
        {
            int original_seg_size = segReader.Count;

            var output = new QueryOutputCells<T>(shapeid, this.Cells.Count, segReader.GetNextSegment(this.Cells.Count));

            int final_seg_size = segReader.Count;

            if ((final_seg_size - original_seg_size) != output.__totalcellcount)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException("Unexpected cursor");
            }

            return output;
        }

        private int _get_total_cell_count(int numshapes)
        {
            // Count the cells not in sections
            int count = this.Cells.Count * numshapes;

            // Count the Cells in the Sections
            return count;
        }

        private Streams.StreamArray _build_src_stream()
        {
            int dummy_shapeid = -1;
            int numshapes = 1;
            int shapeindex = 0;
            int numcells = this._get_total_cell_count(numshapes);
            var stream = new VisioAutomation.ShapeSheet.Streams.FixedSrcStreamBuilder(numcells);
            var cellinfos = this._enum_total_cellinfo(dummy_shapeid, shapeindex);
            var srcs = cellinfos.Select(i => i.SidSrc.Src);
            stream.AddRange(srcs);

            return stream.ToStream();
        }

        private VisioAutomation.ShapeSheet.Streams.StreamArray _build_sidsrc_stream(IList<int> shapeids)
        {
            int numshapes = shapeids.Count;
            int numcells = this._get_total_cell_count(numshapes);

            var stream = new VisioAutomation.ShapeSheet.Streams.FixedSidSrcStreamBuilder(numcells);

            for (int shapeindex = 0; shapeindex < shapeids.Count; shapeindex++)
            {
                // For each shape add the cells to query
                var shapeid = shapeids[shapeindex];

                var cellinfos = this._enum_total_cellinfo(shapeid, shapeindex);
                var sidsrcs = cellinfos.Select(i => i.SidSrc);
                stream.AddRange(sidsrcs);
            }

            return stream.ToStream();
        }

        private IEnumerable<Internal.QueryCellInfo> _enum_total_cellinfo(int shapeid, int shapeindex)
        {
            // enum Cells
            foreach (var col in this.Cells)
            {
                var sidsrc = new SidSrc((short)shapeid, col.Src);

                var cellinfo = new Internal.QueryCellInfo(sidsrc, col);
                yield return cellinfo;
            }
        }
    }
}