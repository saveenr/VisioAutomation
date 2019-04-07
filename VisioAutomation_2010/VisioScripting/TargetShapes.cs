using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioScripting
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

            if (this._items == null)
            {
                throw new System.ArgumentException("Target shapes must be resolved before calling ToShapeIDs()");
            }

            var shapeids = this._items.Select(s => s.ID); 
            var target_shapeids = new List<int>(this._items.Count);
            target_shapeids.AddRange(shapeids);
            return target_shapeids;
        }

        public VisioAutomation.ShapeIDPairs ToShapeIDPairs()
        {
            _verify_resolved();

            if (this._items == null)
            {
                throw new System.ArgumentException("Target shapes must be resolved before calling ToShapeIDs()");
            }

            return VisioAutomation.ShapeIDPairs.FromShapes(this._items);
        }

        internal int SelectShapesAndCount(VisioScripting.Client client)
        {
            client.Application.AssertHasAttachedApplication();

            var app = client.Application.GetAttachedApplication();
            var active_window = app.ActiveWindow;
            var sel = active_window.Selection;

            if (this._items == null)
            {
                int n = sel.Count;
                client.Output.WriteVerbose("GetTargetSelectionCount: Using active selection of {0} shapes", n);
                return n;
            }

            client.Output.WriteVerbose("GetTargetSelectionCount: Resetting selection to specified {0} shapes", this._items.Count);

            // Force empty selection
            active_window.DeselectAll();
            active_window.DeselectAll(); // doing this twice is deliberate

            // Force selection to specific shapes
            active_window.Select(this._items, IVisio.VisSelectArgs.visSelect);

            int selected_count = sel.Count;
            return selected_count;
        }

        public TargetShapes Resolve(VisioScripting.Client client)
        {
            if (this.IsResolved)
            {
                return this;
            }

            var shapes = client.Selection.GetShapes(new VisioScripting.TargetSelection());
            var targetshapes = new TargetShapes(shapes);
            return targetshapes;
        }

        private void _verify_resolved()
        {
            if (!this.IsResolved)
            {
                throw new System.ArgumentException("This method only supported when the target shapes have been resolved");
            }
        }

        public IList<IVisio.Shape> Shapes => this._items;
    }
}