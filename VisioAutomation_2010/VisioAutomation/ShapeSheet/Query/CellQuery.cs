using VA = VisioAutomation;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class QueryEx
    {
        public class SectionSubQuery
        {
            public short SectionIndex { get; private set; }
            public List<QueryColumn> Columns { get; private set; }
            public int Ordinal;

            public SectionSubQuery(int ordinal,short section)
            {
                this.Ordinal = ordinal;
                this.SectionIndex = section;
                this.Columns = new List<QueryColumn>();
            }

            public static implicit operator int(SectionSubQuery m)
            {
                return m.Ordinal;
            }


            public QueryColumn AddColumn(SRC src, string name)
            {
                int ordinal = this.Columns.Count;
                var col=  new QueryColumn(ordinal,src,name);
                this.Columns.Add(col);
                return col;
            }

            public QueryColumn AddColumn(short cell, string name)
            {
                int ordinal = this.Columns.Count;
                var col = new QueryColumn(ordinal, cell, name);
                this.Columns.Add(col);
                return col;
            }
        }

        public class ShapeSectionInfo
        {
            public QueryEx.SectionSubQuery SectionSubQuery { get; private set; }
            public short ShapeID { get; private set; }
            public List<short> RowIndexes { get; private set; }

            public ShapeSectionInfo(SectionSubQuery sq, short shapeid, int numrows)
            {
                this.SectionSubQuery = sq;
                this.ShapeID = shapeid;
                this.RowIndexes = Enumerable.Range(0, numrows).Select(i => (short) i).ToList();
            }
        }

        public List<QueryColumn> Columns { get; private set; }
        public List<SectionSubQuery> Sections { get; private set; }
        private List<List<ShapeSectionInfo>> PerShapeSectionInfo; 

        private bool IsFrozen;
 
        public QueryEx()
        {
            this.Columns = new List<QueryColumn>();
            this.Sections = new List<SectionSubQuery>();
        }


        public QueryColumn AddColumn(SRC src)
        {
            CheckNotFrozen();
            int ordinal = this.Columns.Count;
            var col = new QueryColumn(ordinal, src, null);
            this.Columns.Add(col);
            return col;
        }
        public QueryColumn AddColumn(SRC src,string name)
        {
            CheckNotFrozen();
            int ordinal = this.Columns.Count;
            var col = new QueryColumn(ordinal, src, null);
            this.Columns.Add(col);
            return col;
        }

        
        public SectionSubQuery AddSection(IVisio.VisSectionIndices section)
        {
            CheckNotFrozen();
            int ordinal = this.Sections.Count;
            // Add error checking for section index
            // Add error checking for cell index
            var sec = new SectionSubQuery(ordinal,(short)section);
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

        public ExQueryResult<string> GetFormulas(IVisio.Shape shape)
        {
            this.Freeze();
            var srcstream = BuildSRCStream(shape);
            var values = VA.ShapeSheet.ShapeSheetHelper.GetFormulasU(shape, srcstream);
            var r = new ExQueryResult<string>(shape.ID);
            FillValuesForShape<string>(values, r, 0,0);

            return r;
        }


        public ExQueryResult<T> GetResults<T>(IVisio.Shape shape)
        {
            this.Freeze();
            var srcstream = BuildSRCStream(shape);


            var unitcodes = Enumerable.Range(0, this.GetTotalCellCount(1)).Select(i => IVisio.VisUnitCodes.visNoCast).ToArray();
            var values = VA.ShapeSheet.ShapeSheetHelper.GetResults<T>(shape, srcstream,unitcodes);
            var r = new ExQueryResult<T>(shape.ID16);
            FillValuesForShape<T>(values, r, 0,0);
            return r;
        }

        public ExQueryResult<CellData<T>> GetFormulasAndResults<T>(IVisio.Shape shape)
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

            var r = new ExQueryResult<CellData<T>>(shape.ID16);
            FillValuesForShape<CellData<T>>(combineddata, r, 0, 0);
            return r;
        }

        public List<ExQueryResult<string>> GetFormulas(IVisio.Page page, IList<int>  shapeids)
        {
            this.Freeze();
            var srcstream = BuildSIDSRCStream(page,shapeids);
            var values = VA.ShapeSheet.ShapeSheetHelper.GetFormulasU(page, srcstream);
            var list = FillValuesForMultipleShapes(shapeids, values, srcstream);
            return list;
        }

        public List<ExQueryResult<T>> GetResults<T>(IVisio.Page page, IList<int> shapeids)
        {
            this.Freeze();
            var srcstream = BuildSIDSRCStream(page, shapeids);
            var unitcodes = Enumerable.Range(0, this.GetTotalCellCount(shapeids.Count)).Select(j => IVisio.VisUnitCodes.visNoCast).ToArray();
            var values = VA.ShapeSheet.ShapeSheetHelper.GetResults<T>(page, srcstream, unitcodes);
            var list = FillValuesForMultipleShapes(shapeids, values, srcstream);
            return list;
        }


        public List<ExQueryResult<CellData<T>>> GetFormulasAndResults<T>(IVisio.Page page, IList<int> shapeids)
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

        private List<ExQueryResult<T>> FillValuesForMultipleShapes<T>(IList<int> shapeids, T[] values, short[] srcstream)
        {
            var list = new List<ExQueryResult<T>>();
            int cellcount = 0;
            for (int shape_index = 0; shape_index < shapeids.Count; shape_index++)
            {
                var shapeid = shapeids[shape_index];
                var r = new ExQueryResult<T>(shapeid);
                cellcount = this.FillValuesForShape<T>(values, r, cellcount, shape_index);
                list.Add(r);
            }
            
            if (cellcount*4 != srcstream.Length)
            {
                throw new VA.AutomationException();
            }
            return list;
        }

        private int FillValuesForShape<T>(T[] array, ExQueryResult<T> result, int start, int shape_index)
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
                List<ShapeSectionInfo> sections = this.PerShapeSectionInfo[shape_index];

                foreach (var section in sections)
                {
                    var sd = new SectionResult<T>();
                    sd.Rows = new List<T[]>();

                    result.SectionCells.Add(sd);

                    foreach (var row_index in section.RowIndexes)
                    {
                        var row_values = new T[section.SectionSubQuery.Columns.Count];
                        int num_cols = row_values.Length;
                        for (int c = 0; c < num_cols; c++)
                        {
                            int index = start + cellcount + c;
                            T value = array[index];
                            row_values[c] = value;
                        }
                        sd.Rows.Add(row_values);
                        cellcount += num_cols;
                    }
                }


            }

            return start + cellcount;
        }

        public int GetTotalCellCount(int numshapes)
        {
            int total_cells_from_sections = this.count_cells_from_sections();
            int total = (this.Columns.Count * numshapes) + total_cells_from_sections;
            return total;
        }

        private short[] BuildSRCStream(IVisio.Shape shape)
        {
            this.PerShapeSectionInfo = new List<List<ShapeSectionInfo>>();

            if (this.Sections.Count>0)
            {
                var section_infos = new List<ShapeSectionInfo>();
                foreach (var sec in this.Sections)
                {
                    // Figure out which rows to query
                    int num_rows = shape.RowCount[sec.SectionIndex];
                    var section_info = new ShapeSectionInfo(sec, shape.ID16, num_rows);
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
                        foreach (var col in section.SectionSubQuery.Columns)
                        {
                            cellcount = add_src(srcstream, cellcount, section.SectionSubQuery.SectionIndex, rowindex, col.SRC.Cell);
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
            this.PerShapeSectionInfo = new List<List<ShapeSectionInfo>>();
            if (this.Sections.Count>0)
            {
                var pageshapes = page.Shapes;
                var shapes = shapeids.Select(id=>pageshapes[id]).ToList();

                for (int n = 0; n < shapeids.Count; n++)
                {
                    var shapeid = (short)shapeids[n];
                    var shape = shapes[n];

                    var section_infos = new List<ShapeSectionInfo>();
                    foreach (var sec in this.Sections)
                    {
                        int num_rows = shape.RowCount[sec.SectionIndex];
                        var section_info = new ShapeSectionInfo(sec, shapeid, num_rows);
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
                            foreach (var col in section.SectionSubQuery.Columns)
                            {
                                cellcount = add_sidsrc(srcstream,cellcount,(short)shapeid,section.SectionSubQuery.SectionIndex,rowindex,col.SRC.Cell);
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

        private int count_cells_from_sections()
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
                            foreach (short col in sd.SectionSubQuery.Columns)
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

    public class ExQueryResult<T>
    {
        public int ShapeID;
        public T[] Cells;
        public List<SectionResult<T>> SectionCells; 

        public ExQueryResult(int sid)
        {
            this.ShapeID = sid;
        }    
    }

    public class SectionResult<T>
    {
        public short SectionIndex;
        public List<T[]> Rows;
    }
}