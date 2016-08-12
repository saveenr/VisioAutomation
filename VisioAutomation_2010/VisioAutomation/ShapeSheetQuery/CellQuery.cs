using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheetQuery
{
    public class CellQuery
    {
        public ListColumnSRC Cells { get; }
        public ListSectionSubQuery SectionSubQueries { get; }

        private List<List<SectionSubQueryDetails>> _per_shape_section_info; 
        private bool _is_frozen;

        public CellQuery()
        {
            this.Cells = new ListColumnSRC(0);
            this.SectionSubQueries = new ListSectionSubQuery(0);
            this._per_shape_section_info = new List<List<SectionSubQueryDetails>>(0);
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

        public ColumnSRC AddCell(ShapeSheet.SRC src, string name)
        {
            if (name == null)
            {
                throw new System.ArgumentException("name");
            }

            var col = this.Cells.Add(src, name);
            return col;
        }

        public SectionSubQuery AddSection(IVisio.VisSectionIndices section)
        {
            var col = this.SectionSubQueries.Add(section);
            return col;
        }

        public QueryResult<string> GetFormulas(IVisio.Shape shape)
        {
            this.Freeze();
            var surface = new ShapeSheet.ShapeSheetSurface(shape);
            var srcstream = this.BuildSRCStream(surface);
            var values = surface.GetFormulasU_SRC(srcstream);
            var r = new QueryResult<string>(shape.ID);
            this.FillValuesForShape<string>(values, r, 0,0);

            return r;
        }


        public QueryResult<T> GetResults<T>(IVisio.Shape shape)
        {
            this.Freeze();

            var surface = new ShapeSheet.ShapeSheetSurface(shape);

            var srcstream = this.BuildSRCStream(surface);
            var unitcodes = this.BuildUnitCodeArray(1);
            var values = surface.GetResults_SRC<T>(srcstream,unitcodes);
            var r = new QueryResult<T>(shape.ID);
            this.FillValuesForShape<T>(values, r, 0,0);
            return r;
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

                if (this._per_shape_section_info.Count>0)
                {
                    var per_shape_data = this._per_shape_section_info[i];
                    foreach (var sec in per_shape_data)
                    {
                        foreach (var row_index in sec.RowIndexes)
                        {
                            foreach (var col in sec.SectionSubQuery.Columns)
                            {
                                unitcodes.Add(col.UnitCode);
                            }
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

        public QueryResult<ShapeSheet.CellData<T>> GetCellData<T>(IVisio.Shape shape)
        {
            this.Freeze();

            var surface = new ShapeSheet.ShapeSheetSurface(shape);

            var srcstream = this.BuildSRCStream(surface);
            var unitcodes = this.BuildUnitCodeArray(1);
            var formulas = surface.GetFormulasU_SRC(srcstream);
            var results = surface.GetResults_SRC<T>(srcstream, unitcodes);

            var combineddata = new ShapeSheet.CellData<T>[results.Length];
            for (int i = 0; i < results.Length; i++)
            {
                combineddata[i] = new ShapeSheet.CellData<T>(formulas[i], results[i]);
            }

            var r = new QueryResult<ShapeSheet.CellData<T>>(shape.ID16);
            this.FillValuesForShape<ShapeSheet.CellData<T>>(combineddata, r, 0, 0);
            return r;
        }

        public QueryResultList<string> GetFormulas(IVisio.Page page, IList<int>  shapeids)
        {
            var surface = new ShapeSheet.ShapeSheetSurface(page);
            return this.GetFormulas(surface, shapeids);
        }

        public QueryResultList<string> GetFormulas(ShapeSheet.ShapeSheetSurface surface, IList<int> shapeids)
        {
            this.Freeze();
            var srcstream = this.BuildSIDSRCStream(surface, shapeids);
            var values = surface.GetFormulasU_SIDSRC(srcstream);
            var list = this.FillValuesForMultipleShapes(shapeids, values);
            return list;
        }


        public QueryResultList<T> GetResults<T>(IVisio.Page page, IList<int> shapeids)
        {
            var surface = new ShapeSheet.ShapeSheetSurface(page);
            return this.GetResults<T>(surface, shapeids);
        }

        public QueryResultList<T> GetResults<T>(ShapeSheet.ShapeSheetSurface surface, IList<int> shapeids)
        {
            this.Freeze();
            var srcstream = this.BuildSIDSRCStream(surface, shapeids);
            var unitcodes = this.BuildUnitCodeArray(shapeids.Count);
            var values = surface.GetResults_SIDSRC<T>(srcstream, unitcodes);
            var list = this.FillValuesForMultipleShapes(shapeids, values);
            return list;
        }

        public QueryResultList<ShapeSheet.CellData<T>> GetCellData<T>(IVisio.Page page, IList<int> shapeids)
        {
            var surface = new ShapeSheet.ShapeSheetSurface(page);
            return this.GetCellData<T>(surface, shapeids);
        }

        public QueryResultList<ShapeSheet.CellData<T>> GetCellData<T>(ShapeSheet.ShapeSheetSurface surface, IList<int> shapeids)
        {
            this.Freeze();

            var srcstream = this.BuildSIDSRCStream(surface, shapeids);
            var unitcodes = this.BuildUnitCodeArray(shapeids.Count);
            T[] results = surface.GetResults_SIDSRC<T>(srcstream, unitcodes);
            string[] formulas  = surface.GetFormulasU_SIDSRC(srcstream);

            // Merge the results and formulas
            var combined_data = new ShapeSheet.CellData<T>[results.Length];
            for (int i = 0; i < results.Length; i++)
            {
                combined_data[i] = new ShapeSheet.CellData<T>(formulas[i], results[i]);
            }

            var r = this.FillValuesForMultipleShapes(shapeids, combined_data);
            return r;
        }

        private QueryResultList<T> FillValuesForMultipleShapes<T>(IList<int> shapeids, T[] values)
        {
            var list = new QueryResultList<T>();
            int cellcount = 0;
            for (int shape_index = 0; shape_index < shapeids.Count; shape_index++)
            {
                var shapeid = shapeids[shape_index];
                var data = new QueryResult<T>(shapeid);
                cellcount = this.FillValuesForShape<T>(values, data, cellcount, shape_index);
                list.Add(data);
            }
            
            return list;
        }

        private int FillValuesForShape<T>(T[] array, QueryResult<T> result, int start, int shape_index)
        {
            // First Copy the Cell Values over
            int cellcount = 0;
            var cellarray = new T[this.Cells.Count];
            for (cellcount = 0; cellcount < this.Cells.Count; cellcount++)
            {
                cellarray[cellcount] = array[start+cellcount];
            }

            result.Cells = cellarray;

            // Now copy the Section values over
            if (this._per_shape_section_info.Count > 0)
            {
                var sections = this._per_shape_section_info[shape_index];

                result.Sections = new List<SectionSubQueryResult<T>>(sections.Count);
                foreach (var section in sections)
                {
                    var section_result = new SectionSubQueryResult<T>(section.RowCount);
                    section_result.Column = section.SectionSubQuery;

                    result.Sections.Add(section_result);

                    foreach (int row_index in section.RowIndexes)
                    {
                        var row_values = new T[section.SectionSubQuery.Columns.Count];
                        int num_cols = row_values.Length;
                        for (int c = 0; c < num_cols; c++)
                        {
                            int index = start + cellcount + c;
                            row_values[c] = array[index];
                        }
                        var sec_res_row = new SectionResultRow<T>(row_values);
                        section_result.Rows.Add( sec_res_row );
                        cellcount += num_cols;
                    }
                }
            }
            return start + cellcount;
        }

        private short[] BuildSRCStream(ShapeSheet.ShapeSheetSurface surface)
        {
            if (surface.Target.Shape == null)
            {
                string msg = "Shape must be set in surface not page or master";
                throw new AutomationException(msg);
            }

            this._per_shape_section_info = new List<List<SectionSubQueryDetails>>();

            if (this.SectionSubQueries.Count>0)
            {
                var section_infos = new List<SectionSubQueryDetails>();
                foreach (var sec in this.SectionSubQueries)
                {
                    // Figure out which rows to query
                    int num_rows = surface.Target.Shape.RowCount[(short)sec.SectionIndex];
                    var section_info = new SectionSubQueryDetails(sec, surface.Target.Shape.ID16, num_rows);
                    section_infos.Add(section_info);
                }
                this._per_shape_section_info.Add(section_infos);
            }

            int total = this.GetTotalCellCount(1);

            var stream_builder = new StreamBuilder(3, total);
            
            foreach (var col in this.Cells)
            {
                var src = col.SRC;
                stream_builder.Add(src.Section,src.Row,src.Cell);
            }

            // And then the sections if any exist
            if (this._per_shape_section_info.Count > 0)
            {
                var data_for_shape = this._per_shape_section_info[0];
                foreach (var section in data_for_shape)
                {
                    foreach (int rowindex in section.RowIndexes)
                    {
                        foreach (var col in section.SectionSubQuery.Columns)
                        {
                            stream_builder.Add((short)section.SectionSubQuery.SectionIndex, (short)rowindex, col.CellIndex);
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

        private short[] BuildSIDSRCStream(ShapeSheet.ShapeSheetSurface surface, IList<int> shapeids)
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
                if (this._per_shape_section_info.Count > 0)
                {
                    var data_for_shape = this._per_shape_section_info[i];
                    foreach (var section in data_for_shape)
                    {
                        foreach (int rowindex in section.RowIndexes)
                        {
                            foreach (var col in section.SectionSubQuery.Columns)
                            {
                                stream_builder.Add(
                                    (short)shapeid,
                                    (short)section.SectionSubQuery.SectionIndex,
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


        private void CalculatePerShapeInfo(ShapeSheet.ShapeSheetSurface surface, IList<int> shapeids)
        {
            this._per_shape_section_info = new List<List<SectionSubQueryDetails>>();

            if (this.SectionSubQueries.Count < 1)
            {
                return;
            }

            var pageshapes = surface.Shapes;

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

                var section_infos = new List<SectionSubQueryDetails>(this.SectionSubQueries.Count);
                foreach (var sec in this.SectionSubQueries)
                {
                    int num_rows = GetNumRowsForSection(shape, sec);
                    var section_info = new SectionSubQueryDetails(sec, shapeid, num_rows);
                    section_infos.Add(section_info);
                }
                this._per_shape_section_info.Add(section_infos);
            }

            if (shapeids.Count != this._per_shape_section_info.Count)
            {
                string msg = string.Format("Expected {0} PerShape structs. Actual = {1}", shapeids.Count,
                    this._per_shape_section_info.Count);
                throw new AutomationException(msg);
            }
        }

        private static short GetNumRowsForSection(IVisio.Shape shape, SectionSubQuery sec)
        {
            // For visSectionObject we know the result is always going to be 1
            // so avoid making the call tp RowCount[]
            if (sec.SectionIndex == IVisio.VisSectionIndices.visSectionObject)
            {
                return 1;
            }

            // For all other cases use RowCount[]
            return shape.RowCount[(short)sec.SectionIndex];
        }

        private int GetTotalCellCount(int numshapes)
        {
            // Count the cells not in sections
            int total_cells_not_in_sections = this.Cells.Count * numshapes;

            // Count the Cells in the Sections
            int total_cells_from_sections = 0;
            foreach (var data_for_shape in this._per_shape_section_info)
            {
                foreach (var section_data in data_for_shape)
                {
                    int cells_in_section = section_data.RowCount * section_data.SectionSubQuery.Columns.Count;
                    total_cells_from_sections += cells_in_section;
                }
            }
            
            int total = total_cells_not_in_sections + total_cells_from_sections;
            return total;
        }
    }
}