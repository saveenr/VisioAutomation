using System.Collections.Generic;
using System.Xml.Linq;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.ShapeGeometry
{
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