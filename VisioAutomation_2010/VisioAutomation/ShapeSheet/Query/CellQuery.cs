using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.ShapeSheet.Query
{
    public class CellQuery
    {
        public Columns Columns { get; }

        public CellQuery()
        {
            this.Columns = new Columns();
        }

        public Data.DataRowCollection<string> GetFormulas(IVisio.Shape shape)
        {
            var srcstream = this._build_src_stream();
            var values = shape.GetFormulasU(srcstream);
            var reader = new VisioAutomation.Internal.ArraySegmentEnumerator<string>(values);
            var row = this._shapedata_to_row(shape.ID16, reader);

            var cellqueryresults = new Data.DataRowCollection<string>(1);
            cellqueryresults.Add(row);

            return cellqueryresults;
        }

        public Data.DataRowCollection<TResult> GetResults<TResult>(IVisio.Shape shape)
        {
            var srcstream = this._build_src_stream();
            const object[] unitcodes = null;
            var values = shape.GetResults<TResult>(srcstream, unitcodes);
            var reader = new VisioAutomation.Internal.ArraySegmentEnumerator<TResult>(values);
            var row = this._shapedata_to_row(shape.ID16, reader);


            var cellqueryresults = new Data.DataRowCollection<TResult>(1);
            cellqueryresults.Add(row);

            return cellqueryresults;
        }

        public Data.DataRowCollection<string> GetFormulas(IVisio.Page page, IList<int> shapeids)
        {
            var srcstream = this._build_sidsrc_stream(shapeids);
            var values = page.GetFormulasU(srcstream);
            var reader = new VisioAutomation.Internal.ArraySegmentEnumerator<string>(values);
            var rows = this._shapesid_to_rows(shapeids, reader);

            var cellqueryresults = new Data.DataRowCollection<string>(rows.Count);
            cellqueryresults.AddRange(rows);

            return cellqueryresults;
        }

        public Data.DataRowCollection<TResult> GetResults<TResult>(IVisio.Page page, IList<int> shapeids)
        {
            var srcstream = this._build_sidsrc_stream(shapeids);
            const object[] unitcodes = null;
            var values = page.GetResults<TResult>(srcstream, unitcodes);
            var reader = new VisioAutomation.Internal.ArraySegmentEnumerator<TResult>(values);
            var rows = this._shapesid_to_rows(shapeids, reader);

            var cellqueryresults = new Data.DataRowCollection<TResult>(rows.Count);
            cellqueryresults.AddRange(rows);

            return cellqueryresults;
        }

        private Data.DataRowCollection<T> _shapesid_to_rows<T>(IList<int> shapeids, VisioAutomation.Internal.ArraySegmentEnumerator<T> seg_enumerator)
        {
            var rows = new Data.DataRowCollection<T>(shapeids.Count);
            foreach (int shapeid in shapeids)
            {
                var row = this._shapedata_to_row((short) shapeid, seg_enumerator);
                rows.Add(row);
            }

            return rows;
        }

        private Data.DataRow<T> _shapedata_to_row<T>(short shapeid, VisioAutomation.Internal.ArraySegmentEnumerator<T> seg_enumerator)
        {
            // From the reader, pull as many cells as there are columns
            int numcols = this.Columns.Count;
            int original_seg_size = seg_enumerator.Count;
            var cells = seg_enumerator.GetNextSegment(numcols);

            // verify that nothing strange has happened
            int final_seg_size = seg_enumerator.Count;
            if ((final_seg_size - original_seg_size) != numcols)
            {
                throw new Exceptions.InternalAssertionException("Unexpected cursor");
            }

            var sec_index = IVisio.VisSectionIndices.visSectionInval;
            var row = new Data.DataRow<T>(shapeid, sec_index, cells);
            return row;
        }

        private Streams.StreamArray _build_src_stream()
        {
            int numshapes = 1;
            int numcells = this.Columns.Count * numshapes;
            var srcs = this._enum_srcs();
            var stream = Streams.StreamArray.FromSrc(numcells, srcs);

            return stream;
        }

        private Streams.StreamArray _build_sidsrc_stream(IList<int> shapeids)
        {
            int numshapes = shapeids.Count;
            int numcells = this.Columns.Count * numshapes;
            var sidsrcs = _enum_sidsrcs(shapeids);
            var stream = Streams.StreamArray.FromSidSrc(numcells, sidsrcs);

            return stream;
        }

        private IEnumerable<Core.Src> _enum_srcs()
        {
            foreach (var col in this.Columns)
            {
                yield return col.Src;
            }
        }

        private IEnumerable<Core.SidSrc> _enum_sidsrcs(IList<int> shapeids)
        {
            foreach (var shapeid in shapeids)
            {
                foreach (var col in this.Columns)
                {
                    yield return new Core.SidSrc((short) shapeid, col.Src);
                }
            }
        }
    }
}