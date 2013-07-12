using VA = VisioAutomation;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class QueryEx
    {
        public class SectionQueryEx
        {
            public short SectionIndex;
            public List<short> CellIndexes;

            public SectionQueryEx(short section)
            {
                this.SectionIndex = section;
                this.CellIndexes = new List<short>();
            }
            public SectionQueryEx(short section, IEnumerable<short> cells) :
                this(section)
            {
                this.CellIndexes.AddRange(cells);
            }

            public int AddCell(SRC src)
            {
                int ordinal = this.CellIndexes.Count;
                this.CellIndexes.Add(src.Cell);
                return ordinal;
            }

            public int AddCell(SRC src,string name)
            {
                int ordinal = this.CellIndexes.Count;
                this.CellIndexes.Add(src.Cell);
                return ordinal;
            }

        }

        public class ShapeSectionInfo
        {
            public QueryEx.SectionQueryEx SectionQuery;
            public short ShapeID;
            public List<short> RowIndexes;

            public ShapeSectionInfo(SectionQueryEx sq, short shapeid, int numrows)
            {
                this.SectionQuery = sq;
                this.ShapeID = shapeid;
                this.RowIndexes = Enumerable.Range(0, numrows).Select(i => (short) i).ToList();
            }
        }

        public List<SRC> Cells;
        public List<SectionQueryEx> Sections;
        private List<List<ShapeSectionInfo>> PerShapeSectionInfo; 


        private bool Frozen;
 
        public QueryEx()
        {
            this.Cells = new List<SRC>();
        }

        public int AddCell(SRC src)
        {
            CheckFrozen();
            int ordinal=this.Cells.Count;
            this.Cells.Add(src);
            return ordinal;
        }

        public int AddCell(SRC src, string name)
        {
            return this.AddCell(src);
        }

        public int AddSection(IVisio.VisSectionIndices section, IList<SRC> srcs)
        {
            CheckFrozen();
            if (this.Sections == null)
            {
                this.Sections = new List<SectionQueryEx>();
            }
            int ordinal = this.Sections.Count;

            // Add error checking for section index
            // Add error checking for cell index
            var sec = new SectionQueryEx((short)section, srcs.Select(i => i.Cell));
            this.Sections.Add(sec);
            return ordinal;
        }

        public SectionQueryEx AddSection(IVisio.VisSectionIndices section)
        {
            CheckFrozen();
            if (this.Sections == null)
            {
                this.Sections = new List<SectionQueryEx>();
            }
            int ordinal = this.Sections.Count;

            // Add error checking for section index
            // Add error checking for cell index
            var sec = new SectionQueryEx((short)section);
            this.Sections.Add(sec);
            return sec;
        }

        private void CheckFrozen()
        {
            if (this.Frozen)
            {
                throw new VA.AutomationException("Frozen");
            }
        }

        private void Freeze()
        {
            this.Frozen = true;            
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
            var unitcodes = Enumerable.Range(0, this.Cells.Count).Select(i => IVisio.VisUnitCodes.visNoCast).ToArray();
            var values = VA.ShapeSheet.ShapeSheetHelper.GetResults<T>(shape, srcstream,unitcodes);
            var r = new ExQueryResult<T>(shape.ID16);
            FillValuesForShape<T>(values, r, 0,0);
            return r;
        }

        public ExQueryResult<CellData<T>> GetFormulasAndResults<T>(IVisio.Shape shape)
        {
            this.Freeze();
            var srcstream = BuildSRCStream(shape);
            var unitcodes = Enumerable.Range(0, this.Cells.Count).Select(i => IVisio.VisUnitCodes.visNoCast).ToArray();
            var results = VA.ShapeSheet.ShapeSheetHelper.GetResults<T>(shape, srcstream, unitcodes);
            var formulas = VA.ShapeSheet.ShapeSheetHelper.GetFormulasU(shape, srcstream);

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
            var unitcodes = Enumerable.Range(0, this.Cells.Count * shapeids.Count).Select(j => IVisio.VisUnitCodes.visNoCast).ToArray();
            var values = VA.ShapeSheet.ShapeSheetHelper.GetResults<T>(page, srcstream, unitcodes);
            var list = FillValuesForMultipleShapes(shapeids, values, srcstream);
            return list;
        }

        public List<ExQueryResult<CellData<T>>> GetFormulasAndResults<T>(IVisio.Page page, IList<int> shapeids)
        {
            this.Freeze();
            var srcstream = BuildSIDSRCStream(page, shapeids);
            var unitcodes = Enumerable.Range(0, this.Cells.Count * shapeids.Count).Select(j => IVisio.VisUnitCodes.visNoCast).ToArray();
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
            var cellarray = new T[this.Cells.Count];
            for (cellcount = 0; cellcount < this.Cells.Count; cellcount++)
            {
                cellarray[cellcount] = array[start+cellcount];
            }

            result.Cells = cellarray;

            // Now copy the Section values over
            if (this.PerShapeSectionInfo != null)
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
                        var row_values = new T[section.SectionQuery.CellIndexes.Count];
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

        private short[] BuildSRCStream(IVisio.Shape shape)
        {

            if (this.Sections != null)
            {
                this.PerShapeSectionInfo = new List<List<ShapeSectionInfo>>();
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

            int total_cells_from_sections = this.count_cells_from_sections();
            int total = this.Cells.Count + total_cells_from_sections;


            int cellcount = 0;
            var srcstream = new short[3*total];
            foreach (var src in this.Cells)
            {
                cellcount = this.add_src(srcstream,cellcount,src.Section,src.Row,src.Cell);
            }

            // And then the sections if any exist
            if (this.PerShapeSectionInfo != null)
            {
                var data_for_shape = this.PerShapeSectionInfo[0];
                foreach (var section in data_for_shape)
                {
                    foreach (short rowindex in section.RowIndexes)
                    {
                        foreach (short cell in section.SectionQuery.CellIndexes)
                        {
                            cellcount = add_src(srcstream, cellcount, section.SectionQuery.SectionIndex, rowindex, cell);
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

            if (this.Sections != null)
            {
                this.PerShapeSectionInfo = new List<List<ShapeSectionInfo>>();
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

            int total_cells_from_sections = this.count_cells_from_sections();
            int total = (this.Cells.Count*shapeids.Count) + total_cells_from_sections;

            int cellcount = 0;
            var srcstream = new short[4 * total];

            for (int i = 0; i < shapeids.Count; i++)
            {
                // For each shape add the cells to query
                var shapeid = shapeids[i];
                foreach (var src in this.Cells)
                {
                    cellcount = add_sidsrc(srcstream, cellcount, (short)shapeid, src.Section, src.Row, src.Cell);
                }

                // And then the sections if any exist
                if (this.PerShapeSectionInfo != null)
                {
                    var data_for_shape = this.PerShapeSectionInfo[i];
                    foreach (var section in data_for_shape)
                    {
                        foreach (short rowindex in section.RowIndexes)
                        {
                            foreach (short cell in section.SectionQuery.CellIndexes)
                            {
                                cellcount = add_sidsrc(srcstream,cellcount,(short)shapeid,section.SectionQuery.SectionIndex,rowindex,cell);
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
                            foreach (short cell in sd.SectionQuery.CellIndexes)
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

    public class CellQuery : QueryBase
    {
        public CellQuery() :
            base()
        {
        }


        public QueryColumn AddColumn(SRC src)
        {
            var col = new QueryColumn(this.Columns.Count, src, null);
            this.AddColumn(col);
            return col;
        }

        public QueryColumn AddColumn(SRC src, string name)
        {
            var col = new QueryColumn(this.Columns.Count, src, name);
            this.AddColumn(col);
            return col;
        }

        public VA.ShapeSheet.Data.Table<CellData<T>> GetFormulasAndResults<T>(IVisio.Shape shape)
        {
            var qds = this._Execute<T>(shape, true, true);
            return qds.CreateMergedTable();
        }
        
        public VA.ShapeSheet.Data.Table<string> GetFormulas(IVisio.Shape shape)
        {
            var qds = this._Execute<double>(shape, true, false);
            return qds.Formulas;
        }

        public VA.ShapeSheet.Data.Table<T> GetResults<T>(IVisio.Shape shape)
        {
            var qds = this._Execute<T>(shape, false, true);
            return qds.Results;
        }

        private VA.Internal.QueryResults<T> _Execute<T>(IVisio.Shape shape, bool getformulas, bool getresults)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            VA.ShapeSheet.ShapeSheetHelper.EnforceValidResultType(typeof(T));

            var shapeids = new[] { shape.ID };
            var groupcounts = new[] { 1 };
            int rowcount = shapeids.Count();
            
            // Build the Stream
            var srcs = this.Columns.Select(col => col.SRC).ToList();

            var stream = VA.ShapeSheet.SRC.ToStream(srcs);
            var unitcodes = getresults ? this.CreateUnitCodeArrayForRows(1) : null;
            var formulas = getformulas ? VA.ShapeSheet.ShapeSheetHelper.GetFormulasU(shape, stream) : null;
            var results = getresults ? VA.ShapeSheet.ShapeSheetHelper.GetResults<T>(shape, stream, unitcodes) : null;
            var groups = VA.ShapeSheet.Query.QueryBase.Build(shapeids, groupcounts, rowcount);
            var table = new VA.Internal.QueryResults<T>(formulas, results, shapeids, this.Columns.Count, rowcount, groups);

            return table;
        }

        public VA.ShapeSheet.Data.Table<VA.ShapeSheet.CellData<T>> GetFormulasAndResults<T>(
        IVisio.Page page,
        IList<int> shapeids)
        {
            var table = this._Execute<T>(page, shapeids, true, true);
            return table.CreateMergedTable();
        }

        public VA.ShapeSheet.Data.Table<string> GetFormulas(
            IVisio.Page page,
            IList<int> shapeids)
        {
            var table = this._Execute<double>(page, shapeids, true, false);
            return table.Formulas;
        }

        public VA.ShapeSheet.Data.Table<T> GetResults<T>(
            IVisio.Page page,
            IList<int> shapeids)
        {
            var table = this._Execute<T>(page, shapeids, false, true);
            return table.Results;
        }

        private VA.Internal.QueryResults<T> _Execute<T>(
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

            VA.ShapeSheet.ShapeSheetHelper.EnforceValidResultType(typeof(T));

            var srcs = Columns.Select(i => i.SRC).ToList();         

            var groupcounts = new int[shapeids.Count];
            for (int i = 0; i < shapeids.Count; i++)
            {
                groupcounts[i] = 1;
            }

            int rowcount = shapeids.Count;
            int total_cells = rowcount * this.Columns.Count;

            // Build the Stream
            var sidsrcs = new List<VA.ShapeSheet.SIDSRC>(total_cells);
            foreach (var id in shapeids)
            {
                foreach (var src in srcs)
                {
                    var sidsrc = new VA.ShapeSheet.SIDSRC((short) id, src);
                    sidsrcs.Add(sidsrc);
                }
            }
            var stream = VA.ShapeSheet.SIDSRC.ToStream(sidsrcs);
            var unitcodes = getresults ? CreateUnitCodeArrayForRows(1) : null;
            var formulas = getformulas ? VA.ShapeSheet.ShapeSheetHelper.GetFormulasU(page, stream) : null;
            var results = getresults ? VA.ShapeSheet.ShapeSheetHelper.GetResults<T>(page, stream, unitcodes) : null;
            var groups = VA.ShapeSheet.Query.QueryBase.Build(shapeids, groupcounts, rowcount);
            var table = new VA.Internal.QueryResults<T>(formulas, results, shapeids, this.Columns.Count, rowcount, groups);

            return table;
        }
    }
}