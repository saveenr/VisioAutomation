using VA = VisioAutomation;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class CellQuery : QueryBase<CellQueryColumn>
    {
        public CellQuery() :
            base()
        {
        }

        public CellQueryColumn AddColumn(SRC src)
        {
            var col = new CellQueryColumn(this.Columns.Count, src, null);
            this.AddColumn(col);
            return col;
        }

        public CellQueryColumn AddColumn(SRC src, string name)
        {
            var col = new CellQueryColumn(this.Columns.Count, src, name);
            this.AddColumn(col);
            return col;
        }

        public QueryDataSet<T> GetFormulasAndResults<T>(IVisio.Shape shape)
        {
            var qds = this._Execute<T>(shape, true, true);
            return qds;
        }

        public Table<string> GetFormulas(IVisio.Shape shape)
        {
            var qds = this._Execute<double>(shape, true, false);
            return qds.Formulas;
        }

        public Table<T> GetResults<T>(IVisio.Shape shape)
        {
            var qds = this._Execute<T>(shape, false, true);
            return qds.Results;
        }

        private QueryDataSet<T> _Execute<T>(IVisio.Shape shape, bool getformulas, bool getresults)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }


            int total_cells = Columns.Count;
            var unitcodes = CreateUnitCodeArray();

            if (getresults)
            {
                validate_unitcodes(unitcodes, total_cells);
            }

            var shapeids = new[] { shape.ID };
            var group_counts = new[] { 1 };
            int rowcount = shapeids.Count();
            
            // Build the Stream
            var srcs = this.Columns.Items.Select(col => col.SRC).ToList();
            var stream = SRCStream.FromItems(srcs);

            var formulas = getformulas ? VA.ShapeSheet.ShapeSheetHelper.GetFormulasU(shape, stream) : null;
            var results = getresults ? VA.ShapeSheet.ShapeSheetHelper.GetResults<T>(shape, stream, unitcodes) : null;

            var qds = new QueryDataSet<T>(formulas,results,shapeids,this.Columns.Count, rowcount, group_counts);

            return qds;
        }

        public QueryDataSet<T> GetFormulasAndResults<T>(
                IVisio.Page page,
                IList<int> shapeids)
        {
            var qds = this._Execute<T>(page, shapeids, true, true);
            return qds;
        }

        public Table<string> GetFormulas(
            IVisio.Page page,
            IList<int> shapeids)
        {
            var qds = this._Execute<double>(page, shapeids, true, false);
            return qds.Formulas;
        }

        public Table<T> GetResults<T>(
            IVisio.Page page,
            IList<int> shapeids)
        {
            var qds = this._Execute<T>(page, shapeids, false, true);
            return qds.Results;
        }

        private QueryDataSet<T> _Execute<T>(
            IVisio.Page page,
            IList<int> shapeids, bool getformulas, bool getresults)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            if (shapeids == null)
            {
                throw new System.ArgumentNullException("shapeids");
            }

            var srcs = Columns.Items.Select(i => i.SRC).ToList();
            var unitcodes = CreateUnitCodeArray();
            
            if (getresults)
            {
                validate_unitcodes(unitcodes, srcs.Count);
            }

            var groupcounts = new int[shapeids.Count];
            for (int i = 0; i < shapeids.Count; i++)
            {
                groupcounts[i] = 1;
            }
            int rowcount = shapeids.Count;
            int total_cells = rowcount * this.Columns.Count;

            // Build the Stream
            var stream = new VA.ShapeSheet.SIDSRCStream(total_cells);
            foreach (var id in shapeids)
            {
                foreach (var src in srcs)
                {
                    stream.Add((short) id, src);
                }
            }

            var formulas = getformulas ? VA.ShapeSheet.ShapeSheetHelper.GetFormulasU(page, stream) : null;
            var results = getresults ? VA.ShapeSheet.ShapeSheetHelper.GetResults<T>(page, stream, unitcodes) : null;

            var qds = new QueryDataSet<T>(formulas, results, shapeids, this.Columns.Count, rowcount, groupcounts);

            return qds;
        }
    }
}