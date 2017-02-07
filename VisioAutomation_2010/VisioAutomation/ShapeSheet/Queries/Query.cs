using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Exceptions;
using VisioAutomation.ShapeSheet.Queries.Columns;
using VisioAutomation.ShapeSheet.Queries.Outputs;
using VisioAutomation.ShapeSheet.Utilities;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Queries
{
    public class Query
    {
        public ListColumnQuery Cells { get; }
        public ListSubQuery SubQueries { get; }

        private List<List<SubQuerySectionDetails>> _subquery_shape_info; 

        public Query()
        {
            this.Cells = new ListColumnQuery(0);
            this.SubQueries = new ListSubQuery(0);
            this._subquery_shape_info = new List<List<SubQuerySectionDetails>>(0);
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
            var values = surface.GetFormulasU_SRC(srcstream);
            var shape_index = 0;
            var cursor = 0;
            var subqueryinfo = this._safe_get_subquery_output_for_shape(shape_index);
            var output_for_shape = this._create_output_for_shape<string>(surface.Target.ID16, values, subqueryinfo, ref cursor);

            return output_for_shape;
        }

        public Output<TResult> GetResults<TResult>(ShapeSheetSurface surface)
        {
            var srcstream = this._build_src_stream(surface);
            var unitcodes = this._build_unit_code_array(1);
            var values = surface.GetResults_SRC<TResult>(srcstream, unitcodes);
            var shape_index = 0;
            var cursor = 0;
            var subqueryinfo = this._safe_get_subquery_output_for_shape(shape_index);
            var output_for_shape = this._create_output_for_shape<TResult>(surface.Target.ID16, values, subqueryinfo, ref cursor);
            return output_for_shape;
        }

        public Output<ShapeSheet.CellData> GetFormulasAndResults(ShapeSheetSurface surface)
        {
            var srcstream = this._build_src_stream(surface);
            var unitcodes = this._build_unit_code_array(1);
            var formulas = surface.GetFormulasU_SRC(srcstream);
            var results = surface.GetResults_SRC<string>(srcstream, unitcodes);
            var combined_data = CellData.Combine(formulas, results);

            var shape_index = 0;
            var cursor = 0;
            var subqueryinfo = this._safe_get_subquery_output_for_shape(shape_index);
            var output_for_shape = this._create_output_for_shape<ShapeSheet.CellData>(surface.Target.ID16, combined_data, subqueryinfo, ref cursor);
            return output_for_shape;
        }


        public ListOutput<string> GetFormulas(ShapeSheetSurface surface, IList<int> shapeids)
        {
            var srcstream = this._build_sidsrc_stream(surface, shapeids);
            var values = surface.GetFormulasU_SIDSRC(srcstream);
            var list = this._create_outputs_for_shapes(shapeids, values);
            return list;
        }

        public ListOutput<TResult> GetResults<TResult>(ShapeSheetSurface surface, IList<int> shapeids)
        {
            var srcstream = this._build_sidsrc_stream(surface, shapeids);
            var unitcodes = this._build_unit_code_array(shapeids.Count);
            var values = surface.GetResults_SIDSRC<TResult>(srcstream, unitcodes);
            var list = this._create_outputs_for_shapes(shapeids, values);
            return list;
        }

        public ListOutput<ShapeSheet.CellData> GetFormulasAndResults(ShapeSheetSurface surface, IList<int> shapeids)
        {
            var srcstream = this._build_sidsrc_stream(surface, shapeids);
            var unitcodes = this._build_unit_code_array(shapeids.Count);
            var results = surface.GetResults_SIDSRC<string>(srcstream, unitcodes);
            var formulas  = surface.GetFormulasU_SIDSRC(srcstream);
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
                var subqueryinfo = this._safe_get_subquery_output_for_shape(shape_index);
                var output_for_shape =  this._create_output_for_shape<T>((short)shapeid, values, subqueryinfo, ref cursor);
                output_for_all_shapes.Add(output_for_shape);
            }
            
            return output_for_all_shapes;
        }

        private List<SubQuerySectionDetails> _safe_get_subquery_output_for_shape(int shape_index)
        {
            if (this._subquery_shape_info.Count > 0)
            {
                return this._subquery_shape_info[shape_index];
            }
            return null;
        }

        private Output<T> _create_output_for_shape<T>(short shapeid, T[] values, List<SubQuerySectionDetails> subqueries_details, ref int values_cursor)
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

        private short[] _build_src_stream(ShapeSheetSurface surface)
        {
            if (surface.Target.Shape == null)
            {
                string msg = "Target must be Shape not Page or Master";
                throw new System.ArgumentException(msg);
            }

            this._subquery_shape_info = new List<List<SubQuerySectionDetails>>();

            if (this.SubQueries.Count>0)
            {
                var section_infos = new List<SubQuerySectionDetails>();
                foreach (var sec in this.SubQueries)
                {
                    // Figure out which rows to query
                    int num_rows = surface.Target.Shape.RowCount[(short)sec.SectionIndex];
                    var section_info = new SubQuerySectionDetails(sec, num_rows);
                    section_infos.Add(section_info);
                }
                this._subquery_shape_info.Add(section_infos);
            }

            int total = this._get_total_cell_count(1);

            var stream_builder = new StreamBuilderSRC(total);
            
            foreach (var col in this.Cells)
            {
                var src = col.SRC;
                stream_builder.Add(src.Section,src.Row,src.Cell);
            }

            // And then the sections if any exist
            if (this._subquery_shape_info.Count > 0)
            {
                var data_for_shape = this._subquery_shape_info[0];
                foreach (var section in data_for_shape)
                {
                    foreach (int rowindex in section.RowIndexes)
                    {
                        foreach (var col in section.SubQuery.Columns)
                        {
                            stream_builder.Add((short)section.SubQuery.SectionIndex, (short)rowindex, col.CellIndex);
                        }
                    }
                }
            }

            if (!stream_builder.IsFull)
            {
                string msg = string.Format("StreamBuilder is not full");
                throw new InternalAssertionException(msg);
            }

            return stream_builder.Stream;
        }

        private short[] _build_sidsrc_stream(ShapeSheetSurface surface, IList<int> shapeids)
        {
            this._calcualte_per_shape_info(surface, shapeids);

            int total = this._get_total_cell_count(shapeids.Count);

            var stream_builder = new StreamBuilderSIDSRC(total);

            for (int i = 0; i < shapeids.Count; i++)
            {
                // For each shape add the cells to query
                var shapeid = shapeids[i];
                foreach (var col in this.Cells)
                {
                    var src = col.SRC;
                    stream_builder.Add((short)shapeid, src.Section, src.Row, src.Cell);
                }

                // And then the sections if any exist
                if (this._subquery_shape_info.Count > 0)
                {
                    var data_for_shape = this._subquery_shape_info[i];
                    foreach (var section in data_for_shape)
                    {
                        foreach (int rowindex in section.RowIndexes)
                        {
                            foreach (var col in section.SubQuery.Columns)
                            {
                                stream_builder.Add(
                                    (short)shapeid,
                                    (short)section.SubQuery.SectionIndex,
                                    (short)rowindex,
                                    col.CellIndex);
                            }
                        }
                    }
                }
            }

            if (!stream_builder.IsFull)
            {
                string msg = string.Format("StreamBuilder is not full");
                throw new InternalAssertionException(msg);
            }

            return stream_builder.Stream;
        }


        private void _calcualte_per_shape_info(ShapeSheetSurface surface, IList<int> shapeids)
        {
            this._subquery_shape_info = new List<List<SubQuerySectionDetails>>();

            if (this.SubQueries.Count < 1)
            {
                return;
            }

            var pageshapes = surface.Target.Shapes;

            // For each shapeid fetch the corresponding shape from the page
            // this is needed because we'll need to get per shape section information
            var shapes = new List<IVisio.Shape>(shapeids.Count);
            foreach (int shapeid in shapeids)
            {
                var shape = pageshapes.ItemFromID16[(short)shapeid];
                shapes.Add(shape);
            }

            for (int n = 0; n < shapeids.Count; n++)
            {
                var shape = shapes[n];

                var section_infos = new List<SubQuerySectionDetails>(this.SubQueries.Count);
                foreach (var sec in this.SubQueries)
                {
                    int num_rows = _get_num_rows_for_section(shape, sec);
                    var section_info = new SubQuerySectionDetails(sec, num_rows);
                    section_infos.Add(section_info);
                }
                this._subquery_shape_info.Add(section_infos);
            }

            if (shapeids.Count != this._subquery_shape_info.Count)
            {
                string msg = string.Format("Expected {0} PerShape structs. Actual = {1}", shapeids.Count,
                    this._subquery_shape_info.Count);
                throw new InternalAssertionException(msg);
            }
        }

        private static short _get_num_rows_for_section(IVisio.Shape shape, SubQuery subquery)
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
            foreach (var data_for_shape in this._subquery_shape_info)
            {
                count += data_for_shape.Sum(s => s.RowCount*s.SubQuery.Columns.Count);
            }
            
            return count;
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

                if (this._subquery_shape_info.Count > 0)
                {
                    foreach (var subquery_details in this._subquery_shape_info[i])
                    {
                        foreach (var row_index in subquery_details.RowIndexes)
                        {
                            var subquery_unitcodes = subquery_details.SubQuery.Columns.Select(col => col.UnitCode);
                            unitcodes.AddRange(subquery_unitcodes);
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