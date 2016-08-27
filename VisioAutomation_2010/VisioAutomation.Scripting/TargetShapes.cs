using System.Collections.Generic;
using Microsoft.Office.Interop.Visio;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.Scripting
{
    public class TargetShapes
    {
        public readonly IList<Shape> Shapes;
        public TargetShapes()
        {
            // This explicitly means that the current selection is intended to be used
            this.Shapes = null;
        }
        public TargetShapes(IList<Shape> shapes)
        {
            this.Shapes = shapes;
        }
   
        public int SetSelectionGetSelectedCount(VisioAutomation.Scripting.Client client)
        {
            client.Application.AssertApplicationAvailable();

            if (this.Shapes == null)
            {
                int n = client.Selection.Count();
                client.WriteVerbose("GetTargetSelectionCount: Using active selection of {0} shapes", n);
                return n;
            }

            client.WriteVerbose("GetTargetSelectionCount: Reseting selecton to specified {0} shapes", this.Shapes.Count);
            client.Selection.None();
            client.Selection.Select(this.Shapes);
            int selected_count = client.Selection.Count();
            return selected_count;
        }

        public IList<IVisio.Shape> ResolveShapes(VisioAutomation.Scripting.Client client)
        {
            client.Application.AssertApplicationAvailable();

            if (this.Shapes == null)
            {
                var out_shapes = client.Selection.GetShapes();
                client.WriteVerbose("GetTargetShapes: Returning {0} shapes from the active selection", out_shapes.Count);
                return out_shapes;
            }

            client.WriteVerbose("GetTargetShapes: Returning {0} shapes that were passed in", this.Shapes.Count);
            return this.Shapes;
        }


        public IList<IVisio.Shape> ResolveShapes2DOnly(VisioAutomation.Scripting.Client _client)
        {
            var shapes = this.ResolveShapes(_client);
            var shapes_2d = shapes.Where(s => s.OneD == 0).ToList();
            return shapes_2d;
        }
    }
}