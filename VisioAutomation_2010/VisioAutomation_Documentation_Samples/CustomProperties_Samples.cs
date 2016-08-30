using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation_Documentation_Samples
{
    public static class CustomProperties_Samples
    {
        public static void Set_Custom_Property_on_Shape(IVisio.Document doc)
        {
            // Set Custom Property_on_a_shape

            var page = doc.Pages.Add();
            var s1 = page.DrawRectangle(0, 0, 1, 1);
            var cp = new VisioAutomation.Shapes.CustomProperties.CustomPropertyCells();
            cp.Value = "Hello World";
            VisioAutomation.Shapes.CustomProperties.CustomPropertyHelper.Set(s1, "Propname", cp);

            // Retrieve all the Custom properties from a shape

            var props = VisioAutomation.Shapes.CustomProperties.CustomPropertyHelper.Get(s1);

            // Delete the property from the shape

            VisioAutomation.Shapes.CustomProperties.CustomPropertyHelper.Delete(s1, "Propname");

            //cleanup
            page.Delete(0);
        }

        public static void Set_Custom_Property_on_multiple_Shapes(IVisio.Document doc)
        {
            // Set Custom Property_on_a_shape

            var page = doc.Pages.Add();
            var s1 = page.DrawRectangle(0, 0, 1, 1);
            var s2 = page.DrawRectangle(2, 2, 4, 4);

            var cp1 = new VisioAutomation.Shapes.CustomProperties.CustomPropertyCells();
            cp1.Value = "Hello";
            VisioAutomation.Shapes.CustomProperties.CustomPropertyHelper.Set(s1, "Propname", cp1);

            var cp2 = new VisioAutomation.Shapes.CustomProperties.CustomPropertyCells();
            cp2.Value = "World";
            VisioAutomation.Shapes.CustomProperties.CustomPropertyHelper.Set(s2, "Propname", cp2);

            // Retrieve all the Custom properties from multiple shapes

            var shapes = new[] {s1, s2};
            var props = VisioAutomation.Shapes.CustomProperties.CustomPropertyHelper.Get(page,shapes);

            // Delete the properties from the shapes
            VisioAutomation.Shapes.CustomProperties.CustomPropertyHelper.Delete(s1, "Propname");
            VisioAutomation.Shapes.CustomProperties.CustomPropertyHelper.Delete(s2, "Propname");

            //cleanup
            page.Delete(0);
        }


        public static void Counting_properties(IVisio.Document doc)
        {
            // Set Custom Property_on_a_shape

            var page = doc.Pages.Add();
            var s1 = page.DrawRectangle(0, 0, 1, 1);

            var cp1 = new VisioAutomation.Shapes.CustomProperties.CustomPropertyCells();
            cp1.Value = "Hello";
            VisioAutomation.Shapes.CustomProperties.CustomPropertyHelper.Set(s1, "Propname", cp1);

            int num_custom_props = VisioAutomation.Shapes.CustomProperties.CustomPropertyHelper.GetCount(s1);
            var custom_prop_names = VisioAutomation.Shapes.CustomProperties.CustomPropertyHelper.GetNames(s1);

            //cleanup
            page.Delete(0);
        }
    }
}