using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Geometry
{
    public static class GeometryHelper
    {
        public static short AddGeometrySection(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            int num_geometry_sections = shape.GeometryCount;
            short new_sec_index = GetGeometrySectionIndex((short)num_geometry_sections);
            short actual_sec_index = shape.AddSection(new_sec_index);

            if (actual_sec_index != new_sec_index)
            {
                throw new VA.AutomationException("Internal Error");
            }
            short row_index = shape.AddRow(new_sec_index, (short)IVisio.VisRowIndices.visRowComponent, (short)IVisio.VisRowTags.visTagComponent);

            return new_sec_index;
        }

        private static short GetGeometrySectionIndex(short index)
        {
            short i =
                (short) (((int) IVisio.VisSectionIndices.visSectionFirstComponent) + (index));
            return i;
        }

        public static void DeleteGeometry(IVisio.Shape shape)
        {
            int num = shape.GeometryCount;
            for (int i = num-1; i >=0; i--)
            {
                VA.Geometry.GeometryHelper.DeleteGeometrySection(shape, (short)i);                
            }
        }

        private static void DeleteGeometrySection(IVisio.Shape shape, short index)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            short target_section_index = GetGeometrySectionIndex(index);
            shape.DeleteSection(target_section_index);
        }
    }
}