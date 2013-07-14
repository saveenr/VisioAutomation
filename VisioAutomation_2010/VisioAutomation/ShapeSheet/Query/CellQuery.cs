using VA = VisioAutomation;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public partial class CellQuery
    {
        public List<Column> Columns { get; private set; }
        public List<SectionQuery> Sections { get; private set; }
        private List<List<SectionQueryInfo>> PerShapeSectionInfo; 

        private bool IsFrozen;
 
        public CellQuery()
        {
            this.Columns = new List<Column>();
            this.Sections = new List<SectionQuery>();
        }

        public Column AddColumn(SRC src)
        {
            return this.AddColumn(src, null);
        }

        public Column AddColumn(SRC src,string name)
        {
            CheckNotFrozen();

            if (string.IsNullOrEmpty(name))
            {
                name = string.Format("Col{0}", this.Columns.Count);
            }
            int ordinal = this.Columns.Count;
            var col = new Column(ordinal, src, name);
            this.Columns.Add(col);
            return col;
        }
        
        public SectionQuery AddSection(IVisio.VisSectionIndices section)
        {
            CheckNotFrozen();
            int ordinal = this.Sections.Count;
            // Add error checking for section index
            // Add error checking for cell index
            var sec = new SectionQuery(this,ordinal,(short)section);
            this.Sections.Add(sec);
            return sec;
        }

        private void CheckNotFrozen()
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


            var unitcodes = Enumerable.Range(0, this.GetTotalCellCount(1)).Select(i => IVisio.VisUnitCodes.visNoCast).ToArray();
            var values = VA.ShapeSheet.ShapeSheetHelper.GetResults<T>(shape, srcstream,unitcodes);
            var r = new QueryResult<T>(shape.ID16);
            FillValuesForShape<T>(values, r, 0,0);
            return r;
        }

        public QueryResult<CellData<T>> GetFormulasAndResults<T>(IVisio.Shape shape)
        {
            this.Freeze();

            var srcstream = BuildSRCStream(shape);
            var unitcodes = Enumerable.Range(0, this.GetTotalCellCount(1)).Select(i => IVisio.VisUnitCodes.visNoCast).ToArray();
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

        public QueryResults<string> GetFormulas(IVisio.Page page, IList<int>  shapeids)
        {
            this.Freeze();
            var srcstream = BuildSIDSRCStream(page,shapeids);
            var values = VA.ShapeSheet.ShapeSheetHelper.GetFormulasU(page, srcstream);
            var list = FillValuesForMultipleShapes(shapeids, values, srcstream);
            return list;
        }

        public QueryResults<T> GetResults<T>(IVisio.Page page, IList<int> shapeids)
        {
            this.Freeze();
            var srcstream = BuildSIDSRCStream(page, shapeids);
            var unitcodes = Enumerable.Range(0, this.GetTotalCellCount(shapeids.Count)).Select(j => IVisio.VisUnitCodes.visNoCast).ToArray();
            var values = VA.ShapeSheet.ShapeSheetHelper.GetResults<T>(page, srcstream, unitcodes);
            var list = FillValuesForMultipleShapes(shapeids, values, srcstream);
            return list;
        }


        public QueryResults<CellData<T>> GetFormulasAndResults<T>(IVisio.Page page, IList<int> shapeids)
        {
            this.Freeze();
            var srcstream = BuildSIDSRCStream(page, shapeids);
            var unitcodes = Enumerable.Range(0, this.GetTotalCellCount(shapeids.Count)).Select(j => IVisio.VisUnitCodes.visNoCast).ToArray();
            T[] results = VA.ShapeSheet.ShapeSheetHelper.GetResults<T>(page, srcstream, unitcodes);
            string[] formulas  = VA.ShapeSheet.ShapeSheetHelper.GetFormulasU(page, srcstream);

            var combineddata = new CellData<T>[results.Length];
            for (int i = 0; i < results.Length; i++)
            {
                combineddata[i] = new CellData<T>(formulas[i], results[i]);
            }

            var r = FillValuesForMultipleShapes(shapeids, combineddata, srcstream);
            return r;
        }

        private QueryResults<T> FillValuesForMultipleShapes<T>(IList<int> shapeids, T[] values, short[] srcstream)
        {
            var list = new QueryResults<T>();
            int cellcount = 0;
            for (int shape_index = 0; shape_index < shapeids.Count; shape_index++)
            {
                var shapeid = shapeids[shape_index];
                var r = new QueryResult<T>(shapeid);
                cellcount = this.FillValuesForShape<T>(values, r, cellcount, shape_index);
                list.Add(r);
            }
            
            if (cellcount*4 != srcstream.Length)
            {
                throw new VA.AutomationException();
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
            if (this.PerShapeSectionInfo != null && this.PerShapeSectionInfo.Count > 0)
            {
                result.SectionCells = new List<SectionResult<T>>();
                List<SectionQueryInfo> sections = this.PerShapeSectionInfo[shape_index];

                foreach (var section in sections)
                {
                    var section_result = new SectionResult<T>();
                    section_result.Query = section.SectionQuery;
                    result.SectionCells.Add(section_result);

                    foreach (var row_index in section.RowIndexes)
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

        public int GetTotalCellCount(int numshapes)
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
                    int num_rows = shape.RowCount[sec.SectionIndex];
                    var section_info = new SectionQueryInfo(sec, shape.ID16, num_rows);
                    section_infos.Add(section_info);
                }
                this.PerShapeSectionInfo.Add(section_infos);
            }

            int total = this.GetTotalCellCount(1);
            
            int cellcount = 0;
            var srcstream = new short[3*total];
            foreach (var col in this.Columns)
            {
                var src = col.SRC;
                cellcount = this.add_src(srcstream,cellcount,src.Section,src.Row,src.Cell);
            }

            // And then the sections if any exist
            if (this.PerShapeSectionInfo != null && this.PerShapeSectionInfo.Count > 0)
            {
                var data_for_shape = this.PerShapeSectionInfo[0];
                foreach (var section in data_for_shape)
                {
                    foreach (short rowindex in section.RowIndexes)
                    {
                        foreach (var col in section.SectionQuery.Columns)
                        {
                            cellcount = add_src(srcstream, cellcount, section.SectionQuery.SectionIndex, rowindex, col.SRC.Cell);
                        }
                    }
                }
            }


            if (cellcount != total*3)
            {
                throw new VA.AutomationException();
            }
            return srcstream;
        }

        private short[] BuildSIDSRCStream(IVisio.Page page, IList<int> shapeids)
        {
            this.PerShapeSectionInfo = new List<List<SectionQueryInfo>>();
            if (this.Sections.Count>0)
            {
                var pageshapes = page.Shapes;
                var shapes = shapeids.Select(id=>pageshapes[id]).ToList();

                for (int n = 0; n < shapeids.Count; n++)
                {
                    var shapeid = (short)shapeids[n];
                    var shape = shapes[n];

                    var section_infos = new List<SectionQueryInfo>();
                    foreach (var sec in this.Sections)
                    {
                        int num_rows = shape.RowCount[sec.SectionIndex];
                        var section_info = new SectionQueryInfo(sec, shapeid, num_rows);
                        section_infos.Add(section_info);
                    }
                    this.PerShapeSectionInfo.Add(section_infos);
                }

                if (shapeids.Count != this.PerShapeSectionInfo.Count)
                {
                    throw new VA.AutomationException();
                }                    
            }

            int total = this.GetTotalCellCount(shapeids.Count);

            int cellcount = 0;
            var srcstream = new short[4 * total];

            for (int i = 0; i < shapeids.Count; i++)
            {
                // For each shape add the cells to query
                var shapeid = shapeids[i];
                foreach (var col in this.Columns)
                {
                    var src = col.SRC;
                    cellcount = add_sidsrc(srcstream, cellcount, (short)shapeid, src.Section, src.Row, src.Cell);
                }

                // And then the sections if any exist
                if (this.PerShapeSectionInfo != null && this.PerShapeSectionInfo.Count>0)
                {
                    var data_for_shape = this.PerShapeSectionInfo[i];
                    foreach (var section in data_for_shape)
                    {
                        foreach (short rowindex in section.RowIndexes)
                        {
                            foreach (var col in section.SectionQuery.Columns)
                            {
                                cellcount = add_sidsrc(srcstream,cellcount,(short)shapeid,section.SectionQuery.SectionIndex,rowindex,col.SRC.Cell);
                            }                                
                        }
                    }
                }
            }



            if (cellcount != (total * 4))
            {
                throw new VA.AutomationException();
            }
            return srcstream;
        }

        private int add_sidsrc(short[] srcstream, int i, short id, short section, short row, short cell)
        {
            srcstream[i++] = id;
            srcstream[i++] = section;
            srcstream[i++] = row;
            srcstream[i++] = cell;
            return i;
        }

        private int add_src(short[] srcstream, int i, short section, short row, short cell)
        {
            srcstream[i++] = section;
            srcstream[i++] = row;
            srcstream[i++] = cell;
            return i;
        }

        private int GetCellsCountFromSections()
        {
            int total_cells_from_sections = 0;
            if (this.PerShapeSectionInfo != null)
            {
                foreach (var data_for_shape in this.PerShapeSectionInfo)
                {
                    foreach (var sd in data_for_shape)
                    {
                        foreach (short rowindex in sd.RowIndexes)
                        {
                            foreach (short col in sd.SectionQuery.Columns)
                            {
                                total_cells_from_sections++;
                            } 
                        }
                    }
                }
            }

            return total_cells_from_sections;
        }
    }
}