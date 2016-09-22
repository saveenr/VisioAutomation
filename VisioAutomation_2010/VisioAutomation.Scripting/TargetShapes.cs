using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.Scripting
{
    public class TargetShapes
    {
        public readonly IList<IVisio.Shape> Shapes;
        
        public TargetShapes()
        {
            // This explicitly means that the current selection is intended to be used
            this.Shapes = null;
        }

        public TargetShapeIDs ToShapeIDs(IVisio.Page page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            var shapeids = this.Shapes.Select(s => s.ID).ToList();
            var t = new TargetShapeIDs(page,shapeids);
            return t;
        }

        public TargetShapes(IList<IVisio.Shape> shapes)
        {
            if (shapes == null)
            {
                throw new System.ArgumentNullException(nameof(shapes));
            }
            // If shapes == null then it means to use the active selection
            // else use the specified shapes
            this.Shapes = shapes;
        }

        public TargetShapes(params IVisio.Shape[] shapes)
        {
            if (shapes == null)
            {
                throw new System.ArgumentNullException(nameof(shapes));
            }
            // If shapes == null then it means to use the active selection
            // else use the specified shapes
            this.Shapes = shapes;
        }


        internal int SetSelectionGetSelectedCount(VisioAutomation.Scripting.Client client)
        {
            client.Application.AssertApplicationAvailable();

            if (this.Shapes == null)
            {
                int n = client.Selection.Count();
                client.WriteVerbose("GetTargetSelectionCount: Using active selection of {0} shapes", n);
                return n;
            }

            client.WriteVerbose("GetTargetSelectionCount: Reseting selecton to specified {0} shapes", this.Shapes.Count);
            client.Selection.SelectNone();
            client.Selection.Select(this.Shapes);
            int selected_count = client.Selection.Count();
            return selected_count;
        }

        internal IList<IVisio.Shape> ResolveShapes(VisioAutomation.Scripting.Client client)
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


        internal List<IVisio.Shape> ResolveShapes2DOnly(VisioAutomation.Scripting.Client client)
        {
            var shapes = this.ResolveShapes(client);
            var shapes_2d = shapes.Where(s => s.OneD == 0).ToList();
            return shapes_2d;
        }
    }
}