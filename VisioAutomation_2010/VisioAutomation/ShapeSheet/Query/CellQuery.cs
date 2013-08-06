using VA = VisioAutomation;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public partial class CellQuery
    {
        public ColumnList Columns { get; private set; }
        public SectionList Sections { get; private set; }
        private List<List<SectionQueryInfo>> PerShapeSectionInfo; 
        private bool IsFrozen;

        public CellQuery()
        {
            this.Columns = new ColumnList(0);
            this.Sections = new SectionList(this,0);
            this.PerShapeSectionInfo = new List<List<SectionQueryInfo>>(0);
        }

        internal void CheckNotFrozen()
        {
            if (this.IsFrozen)
            {
                throw new VA.AutomationException("Frozen");
            }
        }

        private void Freeze()
        {
            this.IsFrozen = true;            
        }

        public QueryResult<string> GetFormulas(IVisio.Shape shape)
        {
            this.Freeze();
            var srcstream = BuildSRCStream(shape);
            var values = VA.ShapeSheet.ShapeSheetHelper.GetFormulasU(shape, srcstream);
            var r = new QueryResult<string>(shape.ID);
            FillValuesForShape<string>(values, r, 0,0);

            return r;
        }


        public QueryResult<T> GetResults<T>(IVisio.Shape shape)
        {
            this.Freeze();
            var srcstream = BuildSRCStream(shape);
            var unitcodes = this.BuildUnitCodeArray(1);
            var values = VA.ShapeSheet.ShapeSheetHelper.GetResults<T>(shape, srcstream,unitcodes);
            var r = new QueryResult<T>(shape.ID16);
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
                throw new AutomationException("Internal Error: Number of unit cdes must match number of cells");
            }

            return unitcodes;
        }

        public QueryResult<CellData<T>> GetFormulasAndResults<T>(IVisio.Shape shape)
        {
            this.Freeze();

            var srcstream = BuildSRCStream(shape);
            var unitcodes = this.BuildUnitCodeArray(1);
            var formulas = VA.ShapeSheet.ShapeSheetHelper.GetFormulasU(shape, srcstream);
            var results = VA.ShapeSheet.ShapeSheetHelper.GetResults<T>(shape, srcstream, unitcodes);

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
            this.Freeze();
            var srcstream = BuildSIDSRCStream(page,shapeids);
            var values = VA.ShapeSheet.ShapeSheetHelper.GetFormulasU(page, srcstream);
            var list = FillValuesForMultipleShapes(shapeids, values);
            return list;
        }

        public QueryResultList<T> GetResults<T>(IVisio.Page page, IList<int> shapeids)
        {
            this.Freeze();
            var srcstream = BuildSIDSRCStream(page, shapeids);
            var unitcodes = this.BuildUnitCodeArray(shapeids.Count);
            var values = VA.ShapeSheet.ShapeSheetHelper.GetResults<T>(page, srcstream, unitcodes);
            var list = FillValuesForMultipleShapes(shapeids, values);
            return list;
        }


        public QueryResultList<CellData<T>> GetFormulasAndResults<T>(IVisio.Page page, IList<int> shapeids)
        {
            this.Freeze();
            var srcstream = BuildSIDSRCStream(page, shapeids);
            var unitcodes = this.BuildUnitCodeArray(shapeids.Count);
            T[] results = VA.ShapeSheet.ShapeSheetHelper.GetResults<T>(page, srcstream, unitcodes);
            string[] formulas  = VA.ShapeSheet.ShapeSheetHelper.GetFormulasU(page, srcstream);

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

        private int GetTotalCellCount(int numshapes)
        {
            int total_cells_from_sections = this.GetCellsCountFromSections();
            int total = (this.Columns.Count * numshapes) + total_cells_from_sections;
            return total;
        }

        private short[] BuildSRCStream(IVisio.Shape shape)
        {
            this.PerShapeSectionInfo = new List<List<SectionQueryInfo>>();

            if (this.Sections.Count>0)
            {
                var section_infos = new List<SectionQueryInfo>();
                foreach (var sec in this.Sections)
                {
                    // Figure out which rows to query
                    int num_rows = shape.RowCount[(short)sec.SectionIndex];
                    var section_info = new SectionQueryInfo(sec, shape.ID16, num_rows);
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

        private short[] BuildSIDSRCStream(IVisio.Page page, IList<int> shapeids)
        {
            CalculatePerShapeInfo(page, shapeids);

            int total = this.GetTotalCellCount(shapeids.Count);

            // stream_count is the number of short values that have been written to the array
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
                if (this.PerShapeSectionInfo.Count>0)
                {
                    var data_for_shape = this.PerShapeSectionInfo[i];
                    foreach (var section in data_for_shape)
                    {
                        foreach (int rowindex in section.RowIndexes)
                        {
                            foreach (var col in section.SectionQuery.Columns)
                            {
                                stream_builder.Add((short)shapeid,(short)section.SectionQuery.SectionIndex,(short)rowindex,col.SRC.Cell);
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

        private void CalculatePerShapeInfo(IVisio.Page page, IList<int> shapeids)
        {
            this.PerShapeSectionInfo = new List<List<SectionQueryInfo>>();
            if (this.Sections.Count > 0)
            {
                var pageshapes = page.Shapes;
                var shapes = shapeids.Select(id => pageshapes.ItemFromID16[(short)id]).ToList();

                for (int n = 0; n < shapeids.Count; n++)
                {
                    var shapeid = (short) shapeids[n];
                    var shape = shapes[n];

                    var section_infos = new List<SectionQueryInfo>();
                    foreach (var sec in this.Sections)
                    {
                        int num_rows = shape.RowCount[(short)sec.SectionIndex];
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
        }

        private int GetCellsCountFromSections()
        {
            if (this.PerShapeSectionInfo.Count<1)
            {
                return 0;
            }

            int total_cells_from_sections = 0;
            foreach (var data_for_shape in this.PerShapeSectionInfo)
            {
                foreach (var section_data in data_for_shape)
                {
                    total_cells_from_sections += (section_data.RowCount * section_data.SectionQuery.Columns.Count);
                }
            }

            return total_cells_from_sections;
        }
    }
}