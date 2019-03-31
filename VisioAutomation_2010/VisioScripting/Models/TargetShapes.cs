using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioScripting.Models
{
    public class TargetShapes
    {
        public readonly IList<IVisio.Shape> Shapes;
        
        public TargetShapes()
        {
            // This explicitly means that the current selection is intended to be used
            this.Shapes = null;
        }

        public TargetShapeIDs ToShapeIDs()
        {
            if (this.Shapes == null)
            {
                throw new System.ArgumentException("Target shapes must be resolved before calling ToShapeIDs()");
            }

            var shapeids = this.Shapes.Select(s => s.ID).ToList();
            var target_shapeids = new TargetShapeIDs(shapeids);
            return target_shapeids;
        }

        public VisioAutomation.ShapeIDPairs ToShapeIDPairs()
        {
            if (this.Shapes == null)
            {
                throw new System.ArgumentException("Target shapes must be resolved before calling ToShapeIDs()");
            }

            return VisioAutomation.ShapeIDPairs.FromShapes(this.Shapes);
        }

        public TargetShapes(IList<IVisio.Shape> shapes)
        {
            // If shapes == null then it means to use the active selection
            // else use the specified shapes
            this.Shapes = shapes;
        }

        public TargetShapes(params IVisio.Shape[] shapes)
        {
            // If shapes == null then it means to use the active selection
            // else use the specified shapes
            this.Shapes = shapes;
        }

        internal int SelectShapesAndCount(VisioScripting.Client client)
        {
            client.Application.AssertHasActiveApplication();

            var app = client.Application.GetActiveApplication();
            var active_window = app.ActiveWindow;
            var sel = active_window.Selection;

            if (this.Shapes == null)
            {
                int n = sel.Count;
                client.Output.WriteVerbose("GetTargetSelectionCount: Using active selection of {0} shapes", n);
                return n;
            }

            client.Output.WriteVerbose("GetTargetSelectionCount: Resetting selection to specified {0} shapes", this.Shapes.Count);

            // Force empty slection
            active_window.DeselectAll();
            active_window.DeselectAll(); // doing this twice is deliberate

            // Force selection to specific shapes
            active_window.Select(this.Shapes, IVisio.VisSelectArgs.visSelect);

            int selected_count = sel.Count;
            return selected_count;
        }

        private IList<IVisio.Shape> __ResolveShapes(VisioScripting.Client client)
        {
            client.Application.AssertHasActiveApplication();

            if (this.Shapes == null)
            {
                var out_shapes = client.Selection.GetShapesInSelection();
                client.Output.WriteVerbose("GetTargetShapes: Returning {0} shapes from the active selection", out_shapes.Count);
                return out_shapes;
            }

            client.Output.WriteVerbose("GetTargetShapes: Returning {0} shapes that were passed in", this.Shapes.Count);
            return this.Shapes;
        }

        public TargetShapes  ResolveShapes(VisioScripting.Client client)
        {
            var shapes = this.__ResolveShapes(client);
            var targetshapes = new TargetShapes(shapes);
            return targetshapes;
        }

        internal TargetShapes ResolveShapes2D(VisioScripting.Client client)
        {
            var shapes = this.__ResolveShapes(client);
            var shapes_2d = shapes.Where(s => s.OneD == 0).ToList();
            var targetshapes = new TargetShapes(shapes_2d);
            return targetshapes;
        }
    }
}