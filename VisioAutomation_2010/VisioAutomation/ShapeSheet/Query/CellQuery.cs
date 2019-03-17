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

        public ShapeRow<string> GetFormulas(IVisio.Shape shape)
        {
            var surface = new SurfaceTarget(shape);
            return GetFormulas(surface);
        }


        public ShapeRow<string> GetFormulas(SurfaceTarget surface)
        {
            RestrictToShapesOnly(surface);

            var srcstream = this._build_src_stream();
            var values = surface.GetFormulasU(srcstream);
            var reader = new Internal.ArraySegmentReader<string>(values);
            var output_for_shape = this._create_output_for_shape(surface.ID16, reader);

            return output_for_shape;
        }

        public ShapeRow<TResult> GetResults<TResult>(IVisio.Shape shape)
        {
            var surface = new SurfaceTarget(shape);
            return GetResults<TResult>(surface);
        }

        public ShapeRow<TResult> GetResults<TResult>(SurfaceTarget surface)
        {
            RestrictToShapesOnly(surface);

            var srcstream = this._build_src_stream();
            const object[] unitcodes = null;
            var values = surface.GetResults<TResult>(srcstream, unitcodes);
            var reader = new Internal.ArraySegmentReader<TResult>(values);
            var output_for_shape = this._create_output_for_shape(surface.ID16, reader);
            return output_for_shape;
        }

        public ShapeRowList<string> GetFormulas(IVisio.Page page, IList<int> shapeids)
        {
            var surface = new SurfaceTarget(page);
            return this.GetFormulas(surface, shapeids);
        }


        public ShapeRowList<string> GetFormulas(SurfaceTarget surface, IList<int> shapeids)
        {
            var srcstream = this._build_sidsrc_stream(shapeids);
            var values = surface.GetFormulasU(srcstream);
            var reader = new Internal.ArraySegmentReader<string>(values);
            var list = this._shapesid_to_outputs(shapeids, reader);
            return list;
        }

        public ShapeRowList<TResult> GetResults<TResult>(IVisio.Page page, IList<int> shapeids)
        {
            var surface = new SurfaceTarget(page);
            return this.GetResults<TResult>(surface, shapeids);
        }

        public ShapeRowList<TResult> GetResults<TResult>(SurfaceTarget surface, IList<int> shapeids)
        {
            var srcstream = this._build_sidsrc_stream(shapeids);
            const object[] unitcodes = null;
            var values = surface.GetResults<TResult>(srcstream, unitcodes);
            var reader = new Internal.ArraySegmentReader<TResult>(values);
            var output_list = this._shapesid_to_outputs(shapeids, reader);
            return output_list;
        }

        private ShapeRowList<T> _shapesid_to_outputs<T>(IList<int> shapeids, VASS.Internal.ArraySegmentReader<T> segReader)
        {
            var outputs = shapeids.Select(shapeid => this._create_output_for_shape((short)shapeid, segReader));
            var output_list = new ShapeRowList<T>(shapeids.Count);
            output_list.AddRange(outputs);
            return output_list;
        }

        private ShapeRow<T> _create_output_for_shape<T>(short shapeid, VASS.Internal.ArraySegmentReader<T> segReader)
        {
            // From the reader, pull as many cells as there are columns
            int numcols = this.Columns.Count;
            int original_seg_size = segReader.Count;
            var pulled_segment = segReader.GetNextSegment(numcols);

            // verify that nothing strange has happened
            int final_seg_size = segReader.Count;
            if ((final_seg_size - original_seg_size) != numcols)
            {
                throw new Exceptions.InternalAssertionException("Unexpected cursor");
            }
            
            var output = new ShapeRow<T>(shapeid, numcols, pulled_segment);
            return output;
        }

        private int _get_total_cell_count(int numshapes)
        {
            return this.Columns.Count * numshapes;
        }

        private Streams.StreamArray _build_src_stream()
        {
            int numshapes = 1;
            int numcells = this._get_total_cell_count(numshapes);
            var stream = new VASS.Streams.SrcStreamArrayBuilder(numcells);
            var srcs = this.Columns.Select(c => c.Src);
            stream.AddRange(srcs);

            return stream.ToStreamArray();
        }

        private VASS.Streams.StreamArray _build_sidsrc_stream(IList<int> shapeids)
        {
            int numshapes = shapeids.Count;
            int numcells = this._get_total_cell_count(numshapes);

            var stream = new VASS.Streams.SidSrcStreamArrayBuilder(numcells);
            foreach (var shapeid in shapeids)
            {
                var sidsrcs = this.Columns.Select(c => new SidSrc((short)shapeid, c.Src));
                stream.AddRange(sidsrcs);
            }
            return stream.ToStreamArray();
        }

        private static void RestrictToShapesOnly(SurfaceTarget surface)
        {
            if (surface.Shape == null)
            {
                string msg = "Target must be Shape not Page or Master";
                throw new System.ArgumentException(msg);
            }
        }

    }
}