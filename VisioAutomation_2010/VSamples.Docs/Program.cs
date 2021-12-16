using VSamples.Docs.Samples;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs
{
    class Program
    {
        static void Main(string[] args)
        {

            var app = new IVisio.ApplicationClass();

            var doc = app.Documents.Add("");

            SetCustomProperties.Set_Custom_Property_on_Shape(doc);
            SetCustomProperties.Set_Custom_Property_on_multiple_Shapes(doc);
            DropMasters.One_shape_at_a_time(doc);
            DropMasters.Multiple_shapes_at_a_time(doc);

        }
    }
}
