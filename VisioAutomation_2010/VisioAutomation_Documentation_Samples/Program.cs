using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VisioAutomation_Documentation_Samples
{
    class Program
    {
        static void Main(string[] args)
        {

            var app = new Microsoft.Office.Interop.Visio.ApplicationClass();

            var doc = app.Documents.Add("");

            CustomProperties_Samples.Set_Custom_Property_on_Shape(doc);
            CustomProperties_Samples.Set_Custom_Property_on_multiple_Shapes(doc);
            Dropping_Shapes_Using_Masters.One_shape_at_a_time(doc);
            Dropping_Shapes_Using_Masters.Multiple_shapes_at_a_time(doc);

        }
    }
}
