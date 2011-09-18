using System.Collections.Generic;
using System.Xml.Linq;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.ShapeGeometry
{
    public class GeometrySection
    {
        private List<VA.ShapeGeometry.GeometryRow> Rows; 
        
        public GeometrySection()
        {
            this.Rows = new List<GeometryRow>();
        }
        
        public short AddTo(IVisio.Shape shape)
        {
            short sec_index = ShapeGeometryHelper.AddGeometrySection(shape);
            short row_count = shape.RowCount[sec_index];

            var update = new VA.ShapeSheet.Update.SRCUpdate();
            foreach (var row in this.Rows)
            {
                if (row is GeometryRowMoveTo)
                {
                    var moveto_row = (GeometryRowMoveTo) row;
                    shape.AddRow(sec_index, row_count, (short)IVisio.VisRowTags.visTagMoveTo);
                    var x_src = new VA.ShapeSheet.SRC(sec_index, row_count, VA.ShapeSheet.SRCConstants.Geometry_A.Cell);
                    var y_src = new VA.ShapeSheet.SRC(sec_index, row_count, VA.ShapeSheet.SRCConstants.Geometry_B.Cell);
                    update.SetFormula(x_src, moveto_row.X);
                    update.SetFormula(y_src, moveto_row.Y);
                    row_count++;
                }
                if (row is GeometryRowLineTo)
                {
                    var lineto_row = (GeometryRowLineTo)row;
                    short row_index = shape.AddRow(sec_index, row_count, (short)IVisio.VisRowTags.visTagLineTo);
                    var x_src = new VA.ShapeSheet.SRC(sec_index, row_count, VA.ShapeSheet.SRCConstants.Geometry_A.Cell);
                    var y_src = new VA.ShapeSheet.SRC(sec_index, row_count, VA.ShapeSheet.SRCConstants.Geometry_B.Cell);
                    update.SetFormula(x_src, lineto_row.X);
                    update.SetFormula(y_src, lineto_row.Y);
                    row_count++;
                }
                else
                {
                    
                }
            }

            update.Execute(shape);
            return 0;
        }

        public void MoveTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y)
        {
            var row = new VA.ShapeGeometry.GeometryRowMoveTo(x, y);
            this.Rows.Add(row);
        }

        public void LineTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y)
        {
            var row = new VA.ShapeGeometry.GeometryRowLineTo(x, y);
            this.Rows.Add(row);
        }


    }

    public class GeometryRow
    {
        
    }

    public class GeometryRowMoveTo: GeometryRow
    {
        public VA.ShapeSheet.FormulaLiteral X;
        public VA.ShapeSheet.FormulaLiteral Y;

        internal GeometryRowMoveTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y)
        {
            this.X = x;
            this.Y = y;
        }
    }

    public class GeometryRowLineTo : GeometryRow
    {
        public VA.ShapeSheet.FormulaLiteral X;
        public VA.ShapeSheet.FormulaLiteral Y;

        internal GeometryRowLineTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y)
        {
            this.X = x;
            this.Y = y;
        }
    }

    public static class ShapeGeometryHelper
    {
        public static short AddGeometryMoveToRow(IVisio.Shape shape, short sec_index)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            if (sec_index < ((int)IVisio.VisSectionIndices.visSectionFirstComponent))
            {
                throw new System.ArgumentException("sec_index must be >= VisSectionIndices.visSectionFirstComponent");
            }

            short row_index = shape.AddRow(sec_index, (short)IVisio.VisRowIndices.visRowVertex, (short)IVisio.VisRowTags.visTagMoveTo);


            return row_index;
        }

        public static short AddGeometryLineToRow(IVisio.Shape shape, short sec_index)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            if (sec_index < ((int)IVisio.VisSectionIndices.visSectionFirstComponent))
            {
                throw new System.ArgumentException("sec_index must be >= VisSectionIndices.visSectionFirstComponent");
            }
            short row_index = shape.AddRow(sec_index, (short)IVisio.VisRowIndices.visRowVertex, (short)IVisio.VisRowTags.visTagLineTo);

            return row_index;
        }


        public static short AddGeometrySection(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            int num_geometry_sections = shape.GeometryCount;
            short new_sec_index = (short)(((int)IVisio.VisSectionIndices.visSectionFirstComponent) + (num_geometry_sections));
            short id = shape.AddSection(new_sec_index);
            short row_index = shape.AddRow(new_sec_index, (short)IVisio.VisRowIndices.visRowComponent, (short)IVisio.VisRowTags.visTagComponent);

            return new_sec_index;
        }
    }
}