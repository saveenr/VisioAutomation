
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.ShapeGeometry
{
    public static class ShapeGeometryHelper
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

        public static short GetGeometrySectionIndex(short index)
        {
            short new_sec_index =
                (short) (((int) IVisio.VisSectionIndices.visSectionFirstComponent) + (index));
            return new_sec_index;
        }

        public static void DeleteGeometrySection(IVisio.Shape shape, short section_index)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            int num_geometry_sections = shape.GeometryCount;
            short target_section_index = GetGeometrySectionIndex((short)num_geometry_sections);
            shape.DeleteSection(target_section_index);
        }
    }
}