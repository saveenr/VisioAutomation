using VisioAutomation.Exceptions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public static class ShapeGeometryHelper
    {
        public static short AddSection(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            int num_geometry_sections = shape.GeometryCount;
            short new_sec_index = ShapeGeometryHelper._get_geometry_section_index((short)num_geometry_sections);
            short actual_sec_index = shape.AddSection(new_sec_index);

            if (actual_sec_index != new_sec_index)
            {
                throw new InternalAssertionException();
            }
            short row_index = shape.AddRow(new_sec_index, (short)IVisio.VisRowIndices.visRowComponent, (short)IVisio.VisRowTags.visTagComponent);

            return new_sec_index;
        }

        private static short _get_geometry_section_index(short index)
        {
            short i =
                (short) (((int) IVisio.VisSectionIndices.visSectionFirstComponent) + (index));
            return i;
        }

        public static void Delete(IVisio.Shape shape)
        {
            int num = shape.GeometryCount;
            for (int i = num-1; i >=0; i--)
            {
                ShapeGeometryHelper._delete_section(shape, (short)i);                
            }
        }

        private static void _delete_section(IVisio.Shape shape, short index)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            short target_section_index = ShapeGeometryHelper._get_geometry_section_index(index);
            shape.DeleteSection(target_section_index);
        }
    }
}