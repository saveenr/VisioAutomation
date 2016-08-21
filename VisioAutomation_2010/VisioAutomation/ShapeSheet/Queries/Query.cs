using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Queries.Columns;
using VisioAutomation.ShapeSheet.Queries.Outputs;
using VisioAutomation.ShapeSheet.Queries.Utilities;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Queries
{
    public class Query
    {
        public ListColumnQuery Cells { get; }
        public ListSubQuery SubQueries { get; }

        private List<List<SubQueryDetails>> _subquery_shape_info; 
        private bool _is_frozen;

        public Query()
        {
            this.Cells = new ListColumnQuery(0);
            this.SubQueries = new ListSubQuery(0);
            this._subquery_shape_info = new List<List<SubQueryDetails>>(0);
        }

        internal void CheckNotFrozen()
        {
            if (this._is_frozen)
            {
                throw new AutomationException("Further Modifications to this Query are not allowed");
            }
        }

        private void Freeze()
        {
            this._is_frozen = true;            
        }

        public ColumnQuery AddCell(ShapeSheet.SRC src, string name)
        {
            if (name == null)
            {
                throw new System.ArgumentException("name");
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
            this.Freeze();
            var srcstream = this.BuildSRCStream(surface);
            var values = QueryHelpers.GetFormulasU_SRC(surface, srcstream);
            var shape_index = 0;
            var cursor = 0;
            var output_for_shape = this.CreateOutputForShape<string>(surface.Target.ID16,shape_index, values, ref cursor);

            return output_for_shape;
        }

        public Output<TResult> GetResults<TResult>(ShapeSheetSurface surface)
        {
            this.Freeze();
            var srcstream = this.BuildSRCStream(surface);
            var unitcodes = this.BuildUnitCodeArray(1);
            var values = QueryHelpers.GetResults_SRC<TResult>(surface, srcstream, unitcodes);
            var shape_index = 0;
            var cursor = 0;
            var output_for_shape = this.CreateOutputForShape<TResult>(surface.Target.ID16,shape_index, values, ref cursor);
            return output_for_shape;
        }

        private IList<IVisio.VisUnitCodes> BuildUnitCodeArray(int numshapes)
        {
            if (numshapes < 1)
            {
                throw  new AutomationException("Internal Error: numshapes must be >=1");
            }

            int numcells = this.GetTotalCellCount(numshapes);
            var unitcodes = new List<IVisio.VisUnitCodes>(numcells);

            for (int i = 0; i < numshapes; i++)
            {
                foreach (var col in this.Cells)
                {
                    unitcodes.Add(col.UnitCode);                    
                }

                if (this._subquery_shape_info.Count>0)
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
                throw new AutomationException("Internal Error: Number of unit codes must match number of cells");
            }

            return unitcodes;
        }

        public Output<ShapeSheet.CellData<TResult>> GetFormulasAndResults<TResult>(ShapeSheetSurface surface)
        {
            this.Freeze();

            var srcstream = this.BuildSRCStream(surface);
            var unitcodes = this.BuildUnitCodeArray(1);
            var formulas = QueryHelpers.GetFormulasU_SRC(surface, srcstream);
            var results = QueryHelpers.GetResults_SRC<TResult>(surface, srcstream, unitcodes);
            var combined_data = CellData<TResult>.Combine(results, formulas);

            var shape_index = 0;
            var cursor = 0;
            var output_for_shape = this.CreateOutputForShape<ShapeSheet.CellData<TResult>>(surface.Target.ID16, shape_index, combined_data, ref cursor);
            return output_for_shape;
        }


        public ListOutput<string> GetFormulas(ShapeSheetSurface surface, IList<int> shapeids)
        {
            this.Freeze();
            var srcstream = this.BuildSIDSRCStream(surface, shapeids);
            var values = QueryHelpers.GetFormulasU_SIDSRC(surface, srcstream);
            var list = this.GetOutputsForShapes(shapeids, values);
            return list;
        }

        public ListOutput<TResult> GetResults<TResult>(ShapeSheetSurface surface, IList<int> shapeids)
        {
            this.Freeze();
            var srcstream = this.BuildSIDSRCStream(surface, shapeids);
            var unitcodes = this.BuildUnitCodeArray(shapeids.Count);
            var values = QueryHelpers.GetResults_SIDSRC<TResult>(surface, srcstream, unitcodes);
            var list = this.GetOutputsForShapes(shapeids, values);
            return list;
        }

        public ListOutput<ShapeSheet.CellData<TResult>> GetFormulasAndResults<TResult>(ShapeSheetSurface surface, IList<int> shapeids)
        {
            this.Freeze();

            var srcstream = this.BuildSIDSRCStream(surface, shapeids);
            var unitcodes = this.BuildUnitCodeArray(shapeids.Count);
            var results = QueryHelpers.GetResults_SIDSRC<TResult>(surface, srcstream, unitcodes);
            var formulas  = QueryHelpers.GetFormulasU_SIDSRC(surface, srcstream);
            var combined_data = CellData <TResult>.Combine(results, formulas);
            var r = this.GetOutputsForShapes(shapeids, combined_data);
            return r;
        }

        private ListOutput<T> GetOutputsForShapes<T>(IList<int> shapeids, T[] values)
        {
            var output_for_all_shapes = new ListOutput<T>();

            int cursor = 0;
            for (int shape_index = 0; shape_index < shapeids.Count; shape_index++)
            {
                var shapeid = shapeids[shape_index];
                var output_for_shape =  this.CreateOutputForShape<T>((short)shapeid, shape_index, values, ref cursor );
                output_for_all_shapes.Add(output_for_shape);
            }
            
            return output_for_all_shapes;
        }

        private Output<T> CreateOutputForShape<T>(short shapeid, int shape_index, T[] values, ref int cursor)
        {
            var output = new Output<T>(shapeid);
            // Keep a count of how many cells this method is using
            int cellcount = 0;

            // First Copy the Query Cell Values into the output
            output.Cells = new T[this.Cells.Count];
            for (cellcount = 0; cellcount < this.Cells.Count; cellcount++)
            {
                output.Cells[cellcount] = values[cursor+cellcount];
            }

            // Verify that we have used as many cells as we expect
            if (cellcount != this.Cells.Count)
            {
                throw new VisioAutomation.AutomationException("Internal Error: Mismatch in number of expected cells");                    
            }

            // Now copy the Section values over
            if (this._subquery_shape_info.Count > 0)
            {
                var subqueries = this._subquery_shape_info[shape_index];

                output.Sections = new List<SubQueryOutput<T>>(subqueries.Count);
                foreach (var subquery in subqueries)
                {
                    var subquery_output = new SubQueryOutput<T>(subquery.RowCount);

                    output.Sections.Add(subquery_output);

                    foreach (int row_index in subquery.RowIndexes)
                    {
                        var row_values = new T[subquery.SubQuery.Columns.Count];
                        int num_cols = row_values.Length;
                        for (int c = 0; c < row_values.Length; c++)
                        {
                            int index = cursor + cellcount + c;
                            row_values[c] = values[index];
                        }
                        var sec_res_row = new SubQueryOutputRow<T>(row_values);
                        subquery_output.Rows.Add( sec_res_row );
                        cellcount += num_cols;
                    }
                }
            }

            cursor += cellcount;

            return output;
        }

        private short[] BuildSRCStream(ShapeSheetSurface surface)
        {
            if (surface.Target.Shape == null)
            {
                string msg = "Shape must be set in surface not page or master";
                throw new AutomationException(msg);
            }

            this._subquery_shape_info = new List<List<SubQueryDetails>>();

            if (this.SubQueries.Count>0)
            {
                var section_infos = new List<SubQueryDetails>();
                foreach (var sec in this.SubQueries)
                {
                    // Figure out which rows to query
                    int num_rows = surface.Target.Shape.RowCount[(short)sec.SectionIndex];
                    var section_info = new SubQueryDetails(sec, surface.Target.Shape.ID16, num_rows);
                    section_infos.Add(section_info);
                }
                this._subquery_shape_info.Add(section_infos);
            }

            int total = this.GetTotalCellCount(1);

            var stream_builder = new StreamBuilder(3, total);
            
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

            if (stream_builder.ChunksWrittenCount != total)
            {
                string msg = string.Format("Expected {0} Checks to be written. Actual = {1}", total,
                    stream_builder.ChunksWrittenCount);
                throw new AutomationException(msg);
            }

            return stream_builder.Stream;
        }

        private short[] BuildSIDSRCStream(ShapeSheetSurface surface, IList<int> shapeids)
        {
            this.CalculatePerShapeInfo(surface, shapeids);

            int total = this.GetTotalCellCount(shapeids.Count);

            var stream_builder = new StreamBuilder(4, total);

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

            if (stream_builder.ChunksWrittenCount != total)
            {
                string msg = string.Format("Expected {0} Chunks to be written. Actual = {1}", total,
                    stream_builder.ChunksWrittenCount);
                throw new AutomationException(msg);
            }

            return stream_builder.Stream;
        }


        private void CalculatePerShapeInfo(ShapeSheetSurface surface, IList<int> shapeids)
        {
            this._subquery_shape_info = new List<List<SubQueryDetails>>();

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
                var shapeid = (short)shapeids[n];
                var shape = shapes[n];

                var section_infos = new List<SubQueryDetails>(this.SubQueries.Count);
                foreach (var sec in this.SubQueries)
                {
                    int num_rows = GetNumRowsForSection(shape, sec);
                    var section_info = new SubQueryDetails(sec, shapeid, num_rows);
                    section_infos.Add(section_info);
                }
                this._subquery_shape_info.Add(section_infos);
            }

            if (shapeids.Count != this._subquery_shape_info.Count)
            {
                string msg = string.Format("Expected {0} PerShape structs. Actual = {1}", shapeids.Count,
                    this._subquery_shape_info.Count);
                throw new AutomationException(msg);
            }
        }

        private static short GetNumRowsForSection(IVisio.Shape shape, SubQuery subquery)
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

        private int GetTotalCellCount(int numshapes)
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
    }
}