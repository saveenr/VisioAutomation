using VA = VisioAutomation;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public partial class CellQuery
    {
        public ColumnList Columns { get; private set; }
        public SectionQueryList Sections { get; private set; }

        private List<List<SectionQueryInfo>> PerShapeSectionInfo; 
        private bool IsFrozen;

        public CellQuery()
        {
            this.Columns = new ColumnList(0);
            this.Sections = new SectionQueryList(this,0);
            this.PerShapeSectionInfo = new List<List<SectionQueryInfo>>(0);
        }

        internal void CheckNotFrozen()
        {
            if (this.IsFrozen)
            {
                throw new VA.AutomationException("Further Modifications to this Query are not allowed");
            }
        }

        private void Freeze()
        {
            this.IsFrozen = true;            
        }

        public QueryResult<string> GetFormulas(IVisio.Shape shape)
        {
            this.Freeze();
            var surface = new ShapeSheetSurface(shape);
            var srcstream = BuildSRCStream(surface);
            var values = surface.GetFormulasU_SRC(srcstream);
            var r = new QueryResult<string>(shape.ID);
            FillValuesForShape<string>(values, r, 0,0);

            return r;
        }


        public QueryResult<T> GetResults<T>(IVisio.Shape shape)
        {
            this.Freeze();

            var surface = new ShapeSheetSurface(shape);

            var srcstream = BuildSRCStream(surface);
            var unitcodes = this.BuildUnitCodeArray(1);
            var values = surface.GetResults_SRC<T>(srcstream,unitcodes);
            var r = new QueryResult<T>(shape.ID);
            FillValuesForShape<T>(values, r, 0,0);
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
                foreach (var col in this.Columns)
                {
                    unitcodes.Add(col.UnitCode);                    
                }

                if (this.PerShapeSectionInfo.Count>0)
                {
                    var per_shape_data = this.PerShapeSectionInfo[i];
                    foreach (var sec in per_shape_data)
                    {
                        foreach (var rowindex in sec.RowIndexes)
                        {
                            foreach (var col in sec.SectionQuery.Columns)
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

        public QueryResult<CellData<T>> GetCellData<T>(IVisio.Shape shape)
        {
            this.Freeze();

            var surface = new ShapeSheetSurface(shape);

            var srcstream = BuildSRCStream(surface);
            var unitcodes = this.BuildUnitCodeArray(1);
            var formulas = surface.GetFormulasU_SRC(srcstream);
            var results = surface.GetResults_SRC<T>(srcstream, unitcodes);

            var combineddata = new CellData<T>[results.Length];
            for (int i = 0; i < results.Length; i++)
            {
                combineddata[i] = new CellData<T>(formulas[i], results[i]);
            }

            var r = new QueryResult<CellData<T>>(shape.ID16);
            FillValuesForShape<CellData<T>>(combineddata, r, 0, 0);
            return r;
        }

        public QueryResultList<string> GetFormulas(IVisio.Page page, IList<int>  shapeids)
        {
            var surface = new ShapeSheetSurface(page);
            return this.GetFormulas(surface, shapeids);
        }

        public QueryResultList<string> GetFormulas(ShapeSheetSurface surface, IList<int> shapeids)
        {
            this.Freeze();
            var srcstream = BuildSIDSRCStream(surface, shapeids);
            var values = surface.GetFormulasU_SIDSRC(srcstream);
            var list = FillValuesForMultipleShapes(shapeids, values);
            return list;
        }


        public QueryResultList<T> GetResults<T>(IVisio.Page page, IList<int> shapeids)
        {
            var surface = new ShapeSheetSurface(page);
            return this.GetResults<T>(surface, shapeids);
        }

        public QueryResultList<T> GetResults<T>(ShapeSheetSurface surface, IList<int> shapeids)
        {
            this.Freeze();
            var srcstream = BuildSIDSRCStream(surface, shapeids);
            var unitcodes = this.BuildUnitCodeArray(shapeids.Count);
            var values = surface.GetResults_SIDSRC<T>(srcstream, unitcodes);
            var list = FillValuesForMultipleShapes(shapeids, values);
            return list;
        }

        public QueryResultList<CellData<T>> GetCellData<T>(IVisio.Page page, IList<int> shapeids)
        {
            var surface = new ShapeSheetSurface(page);
            return GetCellData<T>(surface, shapeids);
        }

        public QueryResultList<CellData<T>> GetCellData<T>(ShapeSheetSurface surface, IList<int> shapeids)
        {
            this.Freeze();

            var srcstream = BuildSIDSRCStream(surface, shapeids);
            var unitcodes = this.BuildUnitCodeArray(shapeids.Count);
            T[] results = surface.GetResults_SIDSRC<T>(srcstream, unitcodes);
            string[] formulas  = surface.GetFormulasU_SIDSRC(srcstream);

            // Merge the results and formulas
            var combined_data = new CellData<T>[results.Length];
            for (int i = 0; i < results.Length; i++)
            {
                combined_data[i] = new CellData<T>(formulas[i], results[i]);
            }

            var r = FillValuesForMultipleShapes(shapeids, combined_data);
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
            var cellarray = new T[this.Columns.Count];
            for (cellcount = 0; cellcount < this.Columns.Count; cellcount++)
            {
                cellarray[cellcount] = array[start+cellcount];
            }

            result.Cells = cellarray;

            // Now copy the Section values over
            if (this.PerShapeSectionInfo.Count > 0)
            {
                result.SectionCells = new List<SectionResult<T>>();
                List<SectionQueryInfo> sections = this.PerShapeSectionInfo[shape_index];

                foreach (var section in sections)
                {
                    var section_result = new SectionResult<T>(section.RowCount);
                    section_result.Query = section.SectionQuery;
                    result.SectionCells.Add(section_result);

                    foreach (int row_index in section.RowIndexes)
                    {
                        var row_values = new T[section.SectionQuery.Columns.Count];
                        int num_cols = row_values.Length;
                        for (int c = 0; c < num_cols; c++)
                        {
                            int index = start + cellcount + c;
                            T value = array[index];
                            row_values[c] = value;
                        }
                        section_result.Add(row_values);
                        cellcount += num_cols;
                    }
                }
            }
            return start + cellcount;
        }

        private short[] BuildSRCStream(ShapeSheetSurface surface)
        {
            if (surface.Shape == null)
            {
                string msg = string.Format("Shape must be set in surface not page or master");
                throw new VA.AutomationException(msg);
            }

            this.PerShapeSectionInfo = new List<List<SectionQueryInfo>>();

            if (this.Sections.Count>0)
            {
                var section_infos = new List<SectionQueryInfo>();
                foreach (var sec in this.Sections)
                {
                    // Figure out which rows to query
                    int num_rows = surface.Shape.RowCount[(short)sec.SectionIndex];
                    var section_info = new SectionQueryInfo(sec, surface.Shape.ID16, num_rows);
                    section_infos.Add(section_info);
                }
                this.PerShapeSectionInfo.Add(section_infos);
            }

            int total = this.GetTotalCellCount(1);

            var stream_builder = new StreamBuilder(3, total);
            
            foreach (var col in this.Columns)
            {
                var src = col.SRC;
                stream_builder.Add(src.Section,src.Row,src.Cell);
            }

            // And then the sections if any exist
            if (this.PerShapeSectionInfo.Count > 0)
            {
                var data_for_shape = this.PerShapeSectionInfo[0];
                foreach (var section in data_for_shape)
                {
                    foreach (int rowindex in section.RowIndexes)
                    {
                        foreach (var col in section.SectionQuery.Columns)
                        {
                            stream_builder.Add((short)section.SectionQuery.SectionIndex, (short)rowindex, col.SRC.Cell);
                        }
                    }
                }
            }

            if (stream_builder.ChunksWrittenCount != total)
            {
                string msg = string.Format("Expected {0} Checks to be written. Actual = {1}", total, stream_builder.ChunksWrittenCount);
                throw new VA.AutomationException(msg);
            }

            return stream_builder.Stream;
        }

        private short[] BuildSIDSRCStream(ShapeSheetSurface surface, IList<int> shapeids)
        {
            CalculatePerShapeInfo(surface, shapeids);

            int total = this.GetTotalCellCount(shapeids.Count);

            var stream_builder = new StreamBuilder(4, total);

            for (int i = 0; i < shapeids.Count; i++)
            {
                // For each shape add the cells to query
                var shapeid = shapeids[i];
                foreach (var col in this.Columns)
                {
                    var src = col.SRC;
                    stream_builder.Add((short)shapeid, src.Section, src.Row, src.Cell);
                }

                // And then the sections if any exist
                if (this.PerShapeSectionInfo.Count > 0)
                {
                    var data_for_shape = this.PerShapeSectionInfo[i];
                    foreach (var section in data_for_shape)
                    {
                        foreach (int rowindex in section.RowIndexes)
                        {
                            foreach (var col in section.SectionQuery.Columns)
                            {
                                stream_builder.Add(
                                    (short)shapeid,
                                    (short)section.SectionQuery.SectionIndex,
                                    (short)rowindex,
                                    col.SRC.Cell);
                            }
                        }
                    }
                }
            }

            if (stream_builder.ChunksWrittenCount != total)
            {
                string msg = string.Format("Expected {0} Checks to be written. Actual = {1}", total, stream_builder.ChunksWrittenCount);
                throw new VA.AutomationException(msg);
            }

            return stream_builder.Stream;
        }


        private void CalculatePerShapeInfo(ShapeSheetSurface surface, IList<int> shapeids)
        {
            this.PerShapeSectionInfo = new List<List<SectionQueryInfo>>();

            if (this.Sections.Count < 1)
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

                var section_infos = new List<SectionQueryInfo>(this.Sections.Count);
                foreach (var sec in this.Sections)
                {
                    int num_rows = GetNumRowsForSection(shape, sec);
                    var section_info = new SectionQueryInfo(sec, shapeid, num_rows);
                    section_infos.Add(section_info);
                }
                this.PerShapeSectionInfo.Add(section_infos);
            }

            if (shapeids.Count != this.PerShapeSectionInfo.Count)
            {
                string msg = string.Format("Expected {0} PerShape structs. Actual = {1}", shapeids.Count, this.PerShapeSectionInfo.Count);
                throw new VA.AutomationException(msg);
            }
        }

        private static short GetNumRowsForSection(IVisio.Shape shape, SectionQuery sec)
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
            int total_cells_not_in_sections = this.Columns.Count * numshapes;

            // Count the Cells in the Sections
            int total_cells_from_sections = 0;
            foreach (var data_for_shape in this.PerShapeSectionInfo)
            {
                foreach (var section_data in data_for_shape)
                {
                    int cells_in_section = section_data.RowCount * section_data.SectionQuery.Columns.Count;
                    total_cells_from_sections += cells_in_section;
                }
            }
            
            int total = total_cells_not_in_sections + total_cells_from_sections;
            return total;
        }
    }
}