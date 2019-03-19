using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VASS = VisioAutomation.ShapeSheet;

namespace VisioAutomation.ShapeSheet.Query
{
    public class CellQuery
    {
        public Columns Columns { get; }

        public CellQuery()
        {
            this.Columns = new Columns();
        }

        public CellQueryResults<string> GetFormulas(IVisio.Shape shape)
        {
            var surface = new SurfaceTarget(shape);
            return GetFormulas(surface);
        }


        public CellQueryResults<string> GetFormulas(SurfaceTarget surface)
        {
            RestrictToShapesOnly(surface);

            var srcstream = this._build_src_stream();
            var values = surface.GetFormulasU(srcstream);
            var reader = new Internal.ArraySegmentReader<string>(values);
            var row = this._shapedata_to_row(surface.ID16, reader);

            var cellqueryresults = new CellQueryResults<string>(1);
            cellqueryresults.Add(row);

            return cellqueryresults;
        }

        public CellQueryResults<TResult> GetResults<TResult>(IVisio.Shape shape)
        {
            var surface = new SurfaceTarget(shape);
            return GetResults<TResult>(surface);
        }

        public CellQueryResults<TResult> GetResults<TResult>(SurfaceTarget surface)
        {
            RestrictToShapesOnly(surface);

            var srcstream = this._build_src_stream();
            const object[] unitcodes = null;
            var values = surface.GetResults<TResult>(srcstream, unitcodes);
            var reader = new Internal.ArraySegmentReader<TResult>(values);
            var row = this._shapedata_to_row(surface.ID16, reader);


            var cellqueryresults = new CellQueryResults<TResult>(1);
            cellqueryresults.Add(row);

            return cellqueryresults;
        }

        public CellQueryResults<string> GetFormulas(IVisio.Page page, IList<int> shapeids)
        {
            var surface = new SurfaceTarget(page);
            return this.GetFormulas(surface, shapeids);
        }


        public CellQueryResults<string> GetFormulas(SurfaceTarget surface, IList<int> shapeids)
        {
            var srcstream = this._build_sidsrc_stream(shapeids);
            var values = surface.GetFormulasU(srcstream);
            var reader = new Internal.ArraySegmentReader<string>(values);
            var rows = this._shapesid_to_rows(shapeids, reader);

            var cellqueryresults = new CellQueryResults<string>(rows.Count);
            cellqueryresults.AddRange(rows);

            return cellqueryresults;
        }

        public CellQueryResults<TResult> GetResults<TResult>(IVisio.Page page, IList<int> shapeids)
        {
            var surface = new SurfaceTarget(page);
            return this.GetResults<TResult>(surface, shapeids);
        }

        public CellQueryResults<TResult> GetResults<TResult>(SurfaceTarget surface, IList<int> shapeids)
        {
            var srcstream = this._build_sidsrc_stream(shapeids);
            const object[] unitcodes = null;
            var values = surface.GetResults<TResult>(srcstream, unitcodes);
            var reader = new Internal.ArraySegmentReader<TResult>(values);
            var rows = this._shapesid_to_rows(shapeids, reader);

            var cellqueryresults = new CellQueryResults<TResult>(rows.Count);
            cellqueryresults.AddRange(rows);

            return cellqueryresults;
        }

        private Rows<T> _shapesid_to_rows<T>(IList<int> shapeids, VASS.Internal.ArraySegmentReader<T> segReader)
        {
            var rows = new Rows<T>(shapeids.Count);
            foreach (int shapeid in shapeids)
            {
                var row = this._shapedata_to_row((short)shapeid, segReader);
                rows.Add(row);
            }
            return rows;
        }

        private Row<T> _shapedata_to_row<T>(short shapeid, VASS.Internal.ArraySegmentReader<T> segReader)
        {
            // From the reader, pull as many cells as there are columns
            int numcols = this.Columns.Count;
            int original_seg_size = segReader.Count;
            var cells = segReader.GetNextSegment(numcols);

            // verify that nothing strange has happened
            int final_seg_size = segReader.Count;
            if ((final_seg_size - original_seg_size) != numcols)
            {
                throw new Exceptions.InternalAssertionException("Unexpected cursor");
            }
            
            var row = new Row<T>(shapeid, cells);
            return row;
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