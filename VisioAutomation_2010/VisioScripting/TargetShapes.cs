using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioScripting.Models
{
    public class TargetShapes : TargetObjects<IVisio.Shape>
    {
        
        public TargetShapes() : base()
        {
        }


        public TargetShapes(IList<IVisio.Shape> shapes): base(shapes)
        {
        }

        public TargetShapes(params IVisio.Shape[] shapes) : base(shapes)
        {
        }

        public List<int> ToShapeIDs()
        {
            _verify_resolved();

            if (this.Items == null)
            {
                throw new System.ArgumentException("Target shapes must be resolved before calling ToShapeIDs()");
            }

            var shapeids = this.Items.Select(s => s.ID); 
            var target_shapeids = new List<int>(this.Items.Count);
            target_shapeids.AddRange(shapeids);
            return target_shapeids;
        }

        public VisioAutomation.ShapeIDPairs ToShapeIDPairs()
        {
            _verify_resolved();

            if (this.Items == null)
            {
                throw new System.ArgumentException("Target shapes must be resolved before calling ToShapeIDs()");
            }

            return VisioAutomation.ShapeIDPairs.FromShapes(this.Items);
        }

        internal int SelectShapesAndCount(VisioScripting.Client client)
        {
            client.Application.AssertHasActiveApplication();

            var app = client.Application.GetActiveApplication();
            var active_window = app.ActiveWindow;
            var sel = active_window.Selection;

            if (this.Items == null)
            {
                int n = sel.Count;
                client.Output.WriteVerbose("GetTargetSelectionCount: Using active selection of {0} shapes", n);
                return n;
            }

            client.Output.WriteVerbose("GetTargetSelectionCount: Resetting selection to specified {0} shapes", this.Items.Count);

            // Force empty selection
            active_window.DeselectAll();
            active_window.DeselectAll(); // doing this twice is deliberate

            // Force selection to specific shapes
            active_window.Select(this.Items, IVisio.VisSelectArgs.visSelect);

            int selected_count = sel.Count;
            return selected_count;
        }

        public TargetShapes Resolve(VisioScripting.Client client)
        {
            if (this.IsResolved)
            {
                return this;
            }

            var shapes = client.Selection.GetShapesInSelection();
            var targetshapes = new TargetShapes(shapes);
            return targetshapes;
        }

        internal TargetShapes ResolveShapes2D(VisioScripting.Client client)
        {
            var shapes = client.Selection.GetShapesInSelection();
            var shapes_2d = shapes.Where(s => s.OneD == 0).ToList();
            var targetshapes = new TargetShapes(shapes_2d);
            return targetshapes;
        }

        private void _verify_resolved()
        {
            if (!this.IsResolved)
            {
                throw new System.ArgumentException("This method only supported when the target shapes have been resolved");
            }
        }
    }
}