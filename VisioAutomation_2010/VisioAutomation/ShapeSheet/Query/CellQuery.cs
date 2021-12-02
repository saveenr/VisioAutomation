using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.ShapeSheet.Query
{
    public class CellQuery
    {
        public Data.DataColumnCollection Columns { get; }

        public CellQuery()
        {
            this.Columns = new Data.DataColumnCollection(IVisio.VisSectionIndices.visSectionInval);
        }

        public Data.DataRowCollection<string> GetFormulas(IVisio.Shape shape)
        {
            var srcstream = this._build_src_stream();
            var values = shape.GetFormulasU(srcstream);
            var segments = new VisioAutomation.Internal.ArraySegmentEnumerator<string>(values);
            var row = this._segment_to_row(shape.ID16, segments);

            var datarows = new Data.DataRowCollection<string>(1);
            datarows.Add(row);

            return datarows;
        }

        public Data.DataRowCollection<TResult> GetResults<TResult>(IVisio.Shape shape)
        {
            var srcstream = this._build_src_stream();
            const object[] unitcodes = null;
            var values = shape.GetResults<TResult>(srcstream, unitcodes);
            var segments = new VisioAutomation.Internal.ArraySegmentEnumerator<TResult>(values);
            var row = this._segment_to_row(shape.ID16, segments);


            var datarows = new Data.DataRowCollection<TResult>(1);
            datarows.Add(row);

            return datarows;
        }

        public Data.DataRowCollection<string> GetFormulas(IVisio.Page page, IList<int> shapeids)
        {
            var srcstream = this._build_sidsrc_stream(shapeids);
            var values = page.GetFormulasU(srcstream);
            var segments = new VisioAutomation.Internal.ArraySegmentEnumerator<string>(values);
            var rows = this._get_rows_for_shapes(shapeids, segments);

            var datarows = new Data.DataRowCollection<string>(rows.Count);
            datarows.AddRange(rows);

            return datarows;
        }

        public Data.DataRowCollection<TResult> GetResults<TResult>(IVisio.Page page, IList<int> shapeids)
        {
            var srcstream = this._build_sidsrc_stream(shapeids);
            const object[] unitcodes = null;
            var values = page.GetResults<TResult>(srcstream, unitcodes);
            var segments = new VisioAutomation.Internal.ArraySegmentEnumerator<TResult>(values);
            var rows = this._get_rows_for_shapes(shapeids, segments);

            var datarows = new Data.DataRowCollection<TResult>(rows.Count);
            datarows.AddRange(rows);

            return datarows;
        }

        private Data.DataRowCollection<T> _get_rows_for_shapes<T>(IList<int> shapeids, VisioAutomation.Internal.ArraySegmentEnumerator<T> seg_enumerator)
        {
            var datarows = new Data.DataRowCollection<T>(shapeids.Count);
            foreach (int shapeid in shapeids)
            {
                var row = this._segment_to_row((short) shapeid, seg_enumerator);
                datarows.Add(row);
            }

            return datarows;
        }

        private Data.DataRow<T> _segment_to_row<T>(short shapeid, VisioAutomation.Internal.ArraySegmentEnumerator<T> seg_enumerator)
        {
            // From the reader, pull as many cells as there are columns
            int numcols = this.Columns.Count;
            int original_seg_size = seg_enumerator.Count;
            var segment = seg_enumerator.GetNextSegment(numcols);

            // verify that nothing strange has happened
            int final_seg_size = seg_enumerator.Count;
            if ((final_seg_size - original_seg_size) != numcols)
            {
                throw new Exceptions.InternalAssertionException("Unexpected cursor");
            }

            var sec_index = IVisio.VisSectionIndices.visSectionInval;
            var row = new Data.DataRow<T>(shapeid, sec_index, segment);
            return row;
        }

        private Streams.StreamArray _build_src_stream()
        {
            int numshapes = 1;
            int numcells = this.Columns.Count * numshapes;
            var srcs = this._cols_to_srcs();
            var stream = Streams.StreamArray.FromSrc(numcells, srcs);

            return stream;
        }

        private Streams.StreamArray _build_sidsrc_stream(IList<int> shapeids)
        {
            int numshapes = shapeids.Count;
            int numcells = this.Columns.Count * numshapes;
            var sidsrcs = _cols_to_sidsrcs(shapeids);
            var stream = Streams.StreamArray.FromSidSrc(numcells, sidsrcs);

            return stream;
        }

        private IEnumerable<Core.Src> _cols_to_srcs()
        {
            foreach (var col in this.Columns)
            {
                yield return col.Src;
            }
        }

        private IEnumerable<Core.SidSrc> _cols_to_sidsrcs(IList<int> shapeids)
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