using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Exceptions;
using VisioAutomation.ShapeSheet.Queries.Columns;
using VisioAutomation.ShapeSheet.Queries.Outputs;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Queries
{
    public class Query
    {
        public ListColumnQuery Cells { get; }
        public ListSubQuery SubQueries { get; }

        private List<List<SectionDetails>> _ll_sectiondetails; 

        public Query()
        {
            this.Cells = new ListColumnQuery(0);
            this.SubQueries = new ListSubQuery(0);
            this._ll_sectiondetails = new List<List<SectionDetails>>(0);
        }

        public ColumnQuery AddCell(ShapeSheet.SRC src, string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            var col = this.Cells.Add(src, name);
            return col;
        }

        public SubQuery AddSubQuery(IVisio.VisSectionIndices section)
        {
            var col = this.SubQueries.Add(section);
            return col;
        }

        public Output<string> GetFormulas(ShapeSheetSurface surface)
        {
            var srcstream = this._build_src_stream(surface);
            var values = surface.GetFormulasU(srcstream);
            var shape_index = 0;
            var cursor = 0;
            var subqueryinfo = this.GetSectionDetailsForShape(shape_index);
            var output_for_shape = this._create_output_for_shape<string>(surface.Target.ID16, values, subqueryinfo, ref cursor);

            return output_for_shape;
        }

        public Output<TResult> GetResults<TResult>(ShapeSheetSurface surface)
        {
            var srcstream = this._build_src_stream(surface);
            var unitcodes = this._build_unit_code_array(1);
            var values = surface.GetResults<TResult>(srcstream, unitcodes);
            var shape_index = 0;
            var cursor = 0;
            var subqueryinfo = this.GetSectionDetailsForShape(shape_index);
            var output_for_shape = this._create_output_for_shape<TResult>(surface.Target.ID16, values, subqueryinfo, ref cursor);
            return output_for_shape;
        }

        public Output<ShapeSheet.CellData> GetFormulasAndResults(ShapeSheetSurface surface)
        {
            var srcstream = this._build_src_stream(surface);
            var unitcodes = this._build_unit_code_array(1);
            var formulas = surface.GetFormulasU(srcstream);
            var results = surface.GetResults<string>(srcstream, unitcodes);
            var combined_data = CellData.Combine(formulas, results);

            var shape_index = 0;
            var cursor = 0;
            var subqueryinfo = this.GetSectionDetailsForShape(shape_index);
            var output_for_shape = this._create_output_for_shape<ShapeSheet.CellData>(surface.Target.ID16, combined_data, subqueryinfo, ref cursor);
            return output_for_shape;
        }


        public ListOutput<string> GetFormulas(ShapeSheetSurface surface, IList<int> shapeids)
        {
            this._sidsrc_calculate_per_shape_info(surface, shapeids);
            var srcstream = this._build_sidsrc_stream(surface, shapeids);
            var values = surface.GetFormulasU(srcstream);
            var list = this._create_outputs_for_shapes(shapeids, values);
            return list;
        }

        public ListOutput<TResult> GetResults<TResult>(ShapeSheetSurface surface, IList<int> shapeids)
        {
            this._sidsrc_calculate_per_shape_info(surface, shapeids);
            var srcstream = this._build_sidsrc_stream(surface, shapeids);
            var unitcodes = this._build_unit_code_array(shapeids.Count);
            var values = surface.GetResults<TResult>(srcstream, unitcodes);
            var list = this._create_outputs_for_shapes(shapeids, values);
            return list;
        }

        public ListOutput<ShapeSheet.CellData> GetFormulasAndResults(ShapeSheetSurface surface, IList<int> shapeids)
        {
            this._sidsrc_calculate_per_shape_info(surface, shapeids);
            var srcstream = this._build_sidsrc_stream(surface, shapeids);
            var unitcodes = this._build_unit_code_array(shapeids.Count);
            var results = surface.GetResults<string>(srcstream, unitcodes);
            var formulas  = surface.GetFormulasU(srcstream);
            var combined_data = CellData.Combine(formulas, results);
            var r = this._create_outputs_for_shapes(shapeids, combined_data);
            return r;
        }

        private ListOutput<T> _create_outputs_for_shapes<T>(IList<int> shapeids, T[] values)
        {
            var output_for_all_shapes = new ListOutput<T>();

            int cursor = 0;
            for (int shape_index = 0; shape_index < shapeids.Count; shape_index++)
            {
                var shapeid = shapeids[shape_index];
                var subqueryinfo = this.GetSectionDetailsForShape(shape_index);
                var output_for_shape =  this._create_output_for_shape<T>((short)shapeid, values, subqueryinfo, ref cursor);
                output_for_all_shapes.Add(output_for_shape);
            }
            
            return output_for_all_shapes;
        }

        private List<SectionDetails> GetSectionDetailsForShape(int shape_index)
        {
            if (this._ll_sectiondetails.Count > 0)
            {
                return this._ll_sectiondetails[shape_index];
            }
            return null;
        }

        private Output<T> _create_output_for_shape<T>(short shapeid, T[] values, List<SectionDetails> subqueries_details, ref int values_cursor)
        {
            int old_cursor = values_cursor;

            var output = new Output<T>(shapeid);

            // First Copy the Query Cell Values into the output
            output.Cells = new T[this.Cells.Count];
            for (int i = 0; i < this.Cells.Count; i++)
            {
                output.Cells[i] = values[values_cursor++];
            }

            // Now copy the Section values over
            if (subqueries_details != null)
            {
                output.Sections = new List<SubQueryOutput<T>>(subqueries_details.Count);
                foreach (var subquery_detail in subqueries_details)
                {
                    var subquery_output = new SubQueryOutput<T>(subquery_detail.RowCount);

                    int num_cols = subquery_detail.SubQuery.Columns.Count;
                    foreach (int row_index in subquery_detail.RowIndexes)
                    {
                        var row_values = new T[num_cols];
                        for (int col_index = 0; col_index < num_cols; col_index++)
                        {
                            row_values[col_index] = values[values_cursor++];
                        }
                        var sec_res_row = new SubQueryOutputRow<T>(row_values);
                        subquery_output.Rows.Add(sec_res_row);
                    }

                    output.Sections.Add(subquery_output);

                }
            }

            int num_cells = this.Cells.Count + ( subqueries_details == null ? 0 : subqueries_details.Select( x=>x.RowCount *x.SubQuery.Columns.Count).Sum());
            int expected_cursor = old_cursor + num_cells;
            if (expected_cursor != values_cursor)
            {
                throw new InternalAssertionException("Unexpected cursor");
            }

            return output;
        }

        private void _sidsrc_calculate_per_shape_info(ShapeSheetSurface surface, IList<int> shapeids)
        {
            this._ll_sectiondetails = new List<List<SectionDetails>>();

            if (this.SubQueries.Count < 1)
            {
                return;
            }

            // For each shapeid fetch the corresponding shape from the page
            // this is needed because we'll need to get per shape section information
            var shapes = new List<IVisio.Shape>(shapeids.Count);
            foreach (int shapeid in shapeids)
            {
                var shape = surface.Target.Shapes.ItemFromID16[(short)shapeid];
                shapes.Add(shape);
            }

            for (int n = 0; n < shapeids.Count; n++)
            {
                var shape = shapes[n];

                var l_sectiondetails = new List<SectionDetails>(this.SubQueries.Count);
                foreach (var subquery in this.SubQueries)
                {
                    int num_rows = GetNumRowsForSection(subquery, shape);
                    var sectiondetails = new SectionDetails(subquery, num_rows);
                    l_sectiondetails.Add(sectiondetails);
                }
                this._ll_sectiondetails.Add(l_sectiondetails);
            }

            if (shapeids.Count != this._ll_sectiondetails.Count)
            {
                string msg = string.Format("Expected {0} PerShape structs. Actual = {1}", shapeids.Count,
                    this._ll_sectiondetails.Count);
                throw new InternalAssertionException(msg);
            }
        }

        private static short GetNumRowsForSection(SubQuery subquery, IVisio.Shape shape)
        {
            // For visSectionObject we know the result is always going to be 1
            // so avoid making the call tp RowCount[]
            if (subquery.SectionIndex == IVisio.VisSectionIndices.visSectionObject)
            {
                return 1;
            }

            // For all other cases use RowCount[]
            return shape.RowCount[(short)subquery.SectionIndex];
        }

        private int _get_total_cell_count(int numshapes)
        {
            // Count the cells not in sections
            int count = this.Cells.Count * numshapes;

            // Count the Cells in the Sections
            foreach (var data_for_shape in this._ll_sectiondetails)
            {
                count += data_for_shape.Sum(s => s.RowCount*s.SubQuery.Columns.Count);
            }
            
            return count;
        }



        private short[] _build_src_stream(ShapeSheetSurface surface)
        {
            if (surface.Target.Shape == null)
            {
                string msg = "Target must be Shape not Page or Master";
                throw new System.ArgumentException(msg);
            }

            this._ll_sectiondetails = new List<List<SectionDetails>>();

            if (this.SubQueries.Count > 0)
            {
                var l_sectiondetails = new List<SectionDetails>();
                foreach (var subquery in this.SubQueries)
                {
                    // Figure out which rows to query
                    int num_rows = surface.Target.Shape.RowCount[(short)subquery.SectionIndex];
                    var section_details = new SectionDetails(subquery, num_rows);
                    l_sectiondetails.Add(section_details);
                }
                this._ll_sectiondetails.Add(l_sectiondetails);
            }

            int numshapes = 1;
            int numcells = this._get_total_cell_count(numshapes);
            var streamitem_list = new List<VisioAutomation.ShapeSheet.SRC>(numcells);

            // enum Cells
            foreach (var col in this.Cells)
            {
                var src = col.SRC;
                streamitem_list.Add(src);
            }

            // enum SubQueries
            if (this._ll_sectiondetails.Count > 0)
            {
                var data_for_shape = this._ll_sectiondetails[0];
                foreach (var section in data_for_shape)
                {
                    foreach (int rowindex in section.RowIndexes)
                    {
                        foreach (var col in section.SubQuery.Columns)
                        {
                            var src = new SRC((short)section.SubQuery.SectionIndex, (short)rowindex, col.CellIndex);
                            streamitem_list.Add(src);
                        }
                    }
                }
            }

            if (streamitem_list.Count != numcells)
            {
                string msg = string.Format("src list does not match expected size");
                throw new InternalAssertionException(msg);
            }

            var stream = VisioAutomation.ShapeSheet.SRC.ToStream(streamitem_list);
            return stream;
        }

        private short[] _build_sidsrc_stream(ShapeSheetSurface surface, IList<int> shapeids)
        {
            int numshapes = shapeids.Count;
            int numcells = this._get_total_cell_count(numshapes);

            var streamitem_list = new List<VisioAutomation.ShapeSheet.SIDSRC>(numcells);

            for (int i = 0; i < shapeids.Count; i++)
            {
                // For each shape add the cells to query
                var shapeid = shapeids[i];

                // enum Cells
                foreach (var col in this.Cells)
                {
                    var sidsrc = new SIDSRC((short)shapeid, col.SRC);
                    streamitem_list.Add(sidsrc);
                }

                // enum SubQueries
                if (this._ll_sectiondetails.Count > 0)
                {
                    var data_for_shape = this._ll_sectiondetails[i];
                    foreach (var section in data_for_shape)
                    {
                        foreach (int rowindex in section.RowIndexes)
                        {
                            foreach (var col in section.SubQuery.Columns)
                            {
                                var src = new VisioAutomation.ShapeSheet.SRC(
                                    (short)section.SubQuery.SectionIndex,
                                    (short)rowindex,
                                    col.CellIndex);
                                var sidsrc = new VisioAutomation.ShapeSheet.SIDSRC((short)shapeid, src);
                                streamitem_list.Add(sidsrc);
                            }
                        }
                    }
                }
            }

            var stream = VisioAutomation.ShapeSheet.SIDSRC.ToStream(streamitem_list);

            return stream;
        }

        private List<IVisio.VisUnitCodes> _build_unit_code_array(int numshapes)
        {
            if (numshapes < 1)
            {
                throw new InternalAssertionException("numshapes must be >=1");
            }

            int numcells = this._get_total_cell_count(numshapes);
            var unitcodes = new List<IVisio.VisUnitCodes>(numcells);

            for (int i = 0; i < numshapes; i++)
            {
                foreach (var col in this.Cells)
                {
                    unitcodes.Add(col.UnitCode);
                }

                if (this._ll_sectiondetails.Count > 0)
                {
                    foreach (var data_for_shape in this._ll_sectiondetails[i])
                    {
                        foreach (var row_index in data_for_shape.RowIndexes)
                        {
                            foreach (var col in data_for_shape.SubQuery.Columns)
                            {
                                unitcodes.Add(col.UnitCode);
                            }
                        }
                    }
                }
            }

            if (numcells != unitcodes.Count)
            {
                throw new InternalAssertionException("Number of unit codes must match number of cells");
            }

            return unitcodes;
        }
    }
}