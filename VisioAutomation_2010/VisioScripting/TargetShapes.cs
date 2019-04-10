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
            var items = this._get_items_safe();

            var shapeids = items.Select(s => s.ID); 
            var target_shapeids = new List<int>(items.Count);
            target_shapeids.AddRange(shapeids);
            return target_shapeids;
        }

        public VisioAutomation.ShapeIDPairs ToShapeIDPairs()
        {
            var items = this._get_items_safe();

            return VisioAutomation.ShapeIDPairs.FromShapes(items);
        }

        internal int SelectShapesAndCount(VisioScripting.Client client)
        {
            client.Application.AssertHasAttachedApplication();

            var app = client.Application.GetAttachedApplication();
            var active_window = app.ActiveWindow;
            var sel = active_window.Selection;

            var items = this._get_items_unsafe();
            if (items == null)
            {
                int n = sel.Count;
                client.Output.WriteVerbose("GetTargetSelectionCount: Using active selection of {0} shapes", n);
                return n;
            }

            client.Output.WriteVerbose("GetTargetSelectionCount: Resetting selection to specified {0} shapes", items.Count);

            // Force empty selection
            active_window.DeselectAll();
            active_window.DeselectAll(); // doing this twice is deliberate

            // Force selection to specific shapes
            active_window.Select(items, IVisio.VisSelectArgs.visSelect);

            int selected_count = sel.Count;
            return selected_count;
        }

        public TargetShapes Resolve(VisioScripting.Client client)
        {
            if (this.Resolved)
            {
                return this;
            }

            var target_window = new VisioScripting.TargetWindow();

            var shapes = client.Selection.GetSelectedShapes(target_window);
            var targetshapes = new TargetShapes(shapes);
            return targetshapes;
        }

        public IList<IVisio.Shape> Shapes => this._get_items_safe();
    }
}