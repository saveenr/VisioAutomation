using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs
{
    class Program
    {
        static void Main(string[] args)
        {

            var app = new IVisio.ApplicationClass();

            var doc = app.Documents.Add("");

            Docs.CustomProperties_Samples.Set_Custom_Property_on_Shape(doc);
            Docs.CustomProperties_Samples.Set_Custom_Property_on_multiple_Shapes(doc);
            Docs.Dropping_Shapes_Using_Masters.One_shape_at_a_time(doc);
            Docs.Dropping_Shapes_Using_Masters.Multiple_shapes_at_a_time(doc);

        }
    }
}
