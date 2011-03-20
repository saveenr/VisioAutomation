using System;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionQuery : QueryBase<SectionQueryColumn>
    {
        private readonly short _section;
        private SectionQuery() :
            base()
        {
        }

        public SectionQuery(short section):
            this()
        {
            this._section = section;
        }

        public SectionQuery(IVisio.VisSectionIndices section) :
            this()
        {
            this._section = (short)section;
        }

        public short Section
        {
            get { return _section; }
        }

        public VA.ShapeSheet.SRC GetCellSRCForRow( SectionQueryColumn col, short row)
        {
            var src = new VA.ShapeSheet.SRC(this.Section, row, col.Cell);
            return src;
        }

        public SectionQueryColumn AddColumn(short cell)
        {
            var col = new SectionQueryColumn(this.Columns.Count, cell, null);
            this.AddColumn(col);
            return col;
        }

        public SectionQueryColumn AddColumn(short cell, string name)
        {
            var col = new SectionQueryColumn(this.Columns.Count, cell, name);
            this.AddColumn(col);
            return col;
        }

        public SectionQueryColumn AddColumn(IVisio.VisCellIndices cell)
        {
            return AddColumn((short) cell);
        }

        public SectionQueryColumn AddColumn(IVisio.VisCellIndices cell, string name)
        {
            return AddColumn((short)cell, name);
        }

        private int [] get_group_counts(IVisio.Page page, IList<int> shapeids)
        {
            int num_shapes = shapeids.Count;
            int[] group_counts = new int[num_shapes];
            var page_shapes = page.Shapes;
            for (int i = 0; i < num_shapes; i++)
            {
                int shapeid = shapeids[i];
                var shape = page_shapes.ItemFromID[shapeid];
                group_counts[i] = shape.RowCount[this.Section];
            }
            return group_counts;
        }

        public QueryDataSet<T> GetFormulasAndResults<T>(IVisio.Page page, IList<int> shapeids)
        {
            var qds = this._Execute<T>(page, shapeids,true, true);
            return qds;
        }

        public Table<string> GetFormulas(IVisio.Page page, IList<int> shapeids)
        {
            var qds = this._Execute<double>(page, shapeids, true, true);
            return qds.Formulas;
        }

        public Table<T> GetResults<T>(IVisio.Page page, IList<int> shapeids)
        {
            var qds = this._Execute<T>(page, shapeids, true, true);
            return qds.Results;
        }


        private QueryDataSet<T> _Execute<T>(IVisio.Page page, IList<int> shapeids, bool getformulas, bool getresults)
        {
            if (page == null)
            {
                throw new ArgumentNullException("page");
            }

            if (shapeids == null)
            {
                throw new ArgumentNullException("shapeids");
            }

            var cells = Columns.Items.Select(c => c.Cell).ToList();
            var unitcodes = CreateUnitCodeArray();

            // Find out how many rows are in each shape for the given section id


            // Check preconditions for getting results
            if (getresults)
            {
                validate_unitcodes(unitcodes, cells.Count);
            }

            var groupcounts = this.get_group_counts(page, shapeids);
            var rowcount = groupcounts.Sum();
            int total_cells = rowcount * this.Columns.Count;

            // Build the Stream
            var stream = new VA.ShapeSheet.SIDSRCStream(total_cells);
            for (int shape_index = 0; shape_index < shapeids.Count; shape_index++)
            {
                short shapeid = (short)shapeids[shape_index];
                int num_rows = groupcounts[shape_index];

                for (short row = 0; row < num_rows; row++)
                {
                    foreach (var cell in cells)
                    {
                        stream.Add(shapeid, Section, row, cell);
                    }
                }
            }

            // Retrieve Formulas
            var formulas = getformulas ? VA.ShapeSheet.ShapeSheetHelper.GetFormulasU(page, stream) : null;
            var unitcodes_for_rows = getresults ? ShapeSheetHelper.get_unitcodes_for_rows(unitcodes,rowcount) : null;
            var results = getresults ? VA.ShapeSheet.ShapeSheetHelper.GetResults<T>(page, stream, unitcodes_for_rows) : null;

            var qds = new QueryDataSet<T>(formulas, results, shapeids, this.Columns.Count, rowcount, groupcounts);
            return qds;

        }


        public QueryDataSet<T> GetFormulasAndResults<T>(IVisio.Shape shape)
        {
            var qds =  this._Execute<T>(shape,true,true);
            return qds;
        }

        public Table<string> GetFormulas(IVisio.Shape shape)
        {
            var qds = this._Execute<double>(shape,true,false);
            return qds.Formulas;
        }

        public Table<T> GetResults<T>(IVisio.Shape shape)
        {
            var qds =this._Execute<T>(shape,false,true);
            return qds.Results;
        }

        private QueryDataSet<T> _Execute<T>(IVisio.Shape shape, bool getformulas, bool getresults)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            var cells = Columns.Items.Select(c => c.Cell).ToList();

            int rowcount = shape.RowCount[Section];
            var groupcounts = new[] { rowcount };
            int total_cells = rowcount * Columns.Count;


            var all_unitcodes = getresults ? ShapeSheetHelper.get_unitcodes_for_rows(CreateUnitCodeArray(), rowcount) : null;
            if (getresults)
            {
                validate_unitcodes(all_unitcodes, total_cells);
            }


            // prepare the Stream
            var stream = new VA.ShapeSheet.SRCStream(total_cells);
            for (short row = 0; row < rowcount; row++)
            {
                foreach (var cell in cells)
                {
                    stream.Add(this.Section, row, cell);
                }
            }

            var formulas = getformulas ? VA.ShapeSheet.ShapeSheetHelper.GetFormulasU(shape, stream) : null;
            var results = getresults ? VA.ShapeSheet.ShapeSheetHelper.GetResults<T>(shape, stream, all_unitcodes) : null;

            var shape_ids = new[] { shape.ID };
            var qds = new QueryDataSet<T>(formulas, results, shape_ids, this.Columns.Count, rowcount, groupcounts);

            return qds;
        }
    }
}