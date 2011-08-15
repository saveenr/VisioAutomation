using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Util
{
    public class ShapesheetWalker
    {
        public class SectionInfo
        {
            public IVisio.Shape Shape;
            public int ShapeID;
            public string ShapeNameU;
            public IVisio.VisSectionIndices SectionIndex;
            public string SectionName;
            public bool SectionExists;
            public int RowCount;
            public int MaxCells;
            public short[] CellCountInRows;
        }

        public void Walk(ICollection<IVisio.Shape> shapes)
        {
            this.HandleStartWalk(shapes);
            var map_name_to_secindex = Isotope.EnumUtil.MapNamesToValues<IVisio.VisSectionIndices>();
            map_name_to_secindex.Remove(IVisio.VisSectionIndices.visSectionFirst.ToString());
            map_name_to_secindex.Remove(IVisio.VisSectionIndices.visSectionNone.ToString());
            map_name_to_secindex.Remove(IVisio.VisSectionIndices.visSectionInval.ToString());
            map_name_to_secindex.Remove(IVisio.VisSectionIndices.visSectionLast.ToString());

            foreach (var shape in shapes)
            {
                this.HandleStartShape(shape);

                foreach (var pair_name_to_secindex in map_name_to_secindex)
                {
                    var section_info = new SectionInfo();
                    section_info.ShapeID = shape.ID;
                    section_info.Shape = shape;
                    section_info.ShapeNameU = shape.NameU;
                    section_info.SectionIndex = pair_name_to_secindex.Value;
                    section_info.SectionName = pair_name_to_secindex.Key;
                    section_info.SectionExists = shape.SectionExists[(short)pair_name_to_secindex.Value, 1] != 0;

                    var secindex = (short)pair_name_to_secindex.Value;

                    if (!section_info.SectionExists)
                    {
                        section_info.MaxCells = 0;
                        section_info.RowCount = 0;
                        section_info.SectionExists = false;
                        this.HandleStartSection(section_info);
                        this.HandleEndSection(section_info);
                    }
                    else
                    {
                        section_info.RowCount = shape.RowCount[secindex];
                        section_info.CellCountInRows = Enumerable.Range(0, section_info.RowCount)
                            .Select(r => shape.RowsCellCount[secindex, (short)r])
                            .ToArray();
                        section_info.MaxCells = section_info.CellCountInRows.Max();
                        section_info.SectionExists = true;

                        this.HandleStartSection(section_info);
                        for (int r = 0; r < section_info.RowCount; r++)
                        {
                            int numcells = section_info.CellCountInRows[r];
                            this.HandleStartRow(section_info, (short)r, numcells);

                            for (int c = 0; c < numcells; c++)
                            {
                                var cellsrc = new VisioAutomation.CellSRC(secindex, (short)r, (short)c);
                                this.HandleCell(section_info, cellsrc);
                            }

                            this.HandleEndRow(section_info, (short)r);
                        }
                        this.HandleEndSection(section_info);
                    }
                }
                this.HandleEndShape(shape);
            }
            this.HandleEndWalk();
        }

        public delegate void CellHandler(SectionInfo section_info, VisioAutomation.CellSRC cellsrc);
        public delegate void ShapeHandler(IVisio.Shape shape);
        public delegate void SectionHandler(SectionInfo section_info);

        public event CellHandler OnHandleCell;
        public event ShapeHandler OnEnterShape;
        public event ShapeHandler OnExitShape;
        public event SectionHandler OnEnterSection;
        public event SectionHandler OnExitSection;

        public virtual void HandleStartWalk(ICollection<IVisio.Shape> shapes)
        {
        }

        public virtual void HandleEndWalk()
        {
        }

        public virtual void HandleStartShape(IVisio.Shape shape)
        {
            if (this.OnEnterShape!=null)
            {
                this.OnEnterShape(shape);
            }
        }

        public virtual void HandleEndShape(IVisio.Shape shape)
        {
            if (this.OnExitShape != null)
            {
                this.OnExitShape(shape);
            }
        }

        public virtual void HandleStartSection(SectionInfo section_info)
        {
            if (this.OnEnterSection != null)
            {
                this.OnEnterSection(section_info);
            }
        }

        public virtual void HandleEndSection(SectionInfo section_info)
        {
            if (this.OnExitSection != null)
            {
                this.OnExitSection(section_info);
            }
        }

        public virtual void HandleStartRow(SectionInfo section_info, short rowindex, int numcells)
        {
        }

        public virtual void HandleEndRow(SectionInfo section_info, short rowindex)
        {
        }

        public virtual void HandleCell(SectionInfo section_info, VisioAutomation.CellSRC cellsrc)
        {
            if (this.OnHandleCell != null)
            {
                this.OnHandleCell(section_info, cellsrc);
            }
        }
    }

}
