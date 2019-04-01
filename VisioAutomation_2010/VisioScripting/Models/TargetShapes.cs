using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioScripting.Models
{
    public enum TargetShapesType
    {
        UseSelection,
        UseList
    }
    public class TargetShapes
    {
        private readonly IList<IVisio.Shape> _shapes;
        public readonly TargetShapesType Type;
        
        public TargetShapes()
        {
            // This explicitly means that the current selection is intended to be used
            this._shapes = null;
            this.Type = TargetShapesType.UseSelection;
        }


        public TargetShapes(IList<IVisio.Shape> shapes)
        {
            // If shapes == null then it means to use the active selection
            // else use the specified shapes
            this._shapes = shapes;
            this.Type = this._shapes == null ? TargetShapesType.UseSelection : TargetShapesType.UseList;
        }

        public bool IsResolved
        {
            get
            {
                return (this.Type == TargetShapesType.UseList);
            }
        }
        public int Count
        {
            get
            {
                if (this.Type == TargetShapesType.UseSelection)
                {
                    throw new System.ArgumentException("This method only supported when the target shapes have been resolved");
                }
                return this._shapes.Count;
            }
        }

        public IVisio.Shape this[int index]
        {
            get
            {
                if (this.Type == TargetShapesType.UseSelection)
                {
                    throw new System.ArgumentException("This method only supported when the target shapes have been resolved");
                }
                return this._shapes[index];
            }
        }
        public IList<IVisio.Shape> Shapes
        {
            get
            {
                if (this.Type == TargetShapesType.UseSelection)
                {
                    throw new System.ArgumentException("This method only supported when the target shapes have been resolved");
                }
                return this._shapes;
            }
        }

        public TargetShapes(params IVisio.Shape[] shapes)
        {
            // If shapes == null then it means to use the active selection
            // else use the specified shapes
            this._shapes = shapes;
            this.Type = this._shapes == null ? TargetShapesType.UseSelection : TargetShapesType.UseList;
        }
        public TargetShapeIDs ToShapeIDs()
        {
            if (this.Type == TargetShapesType.UseSelection)
            {
                throw new System.ArgumentException("This method only supported when the target shapes have been resolved");

            }

            if (this._shapes == null)
            {
                throw new System.ArgumentException("Target shapes must be resolved before calling ToShapeIDs()");
            }

            var shapeids = this._shapes.Select(s => s.ID); 
            var target_shapeids = new TargetShapeIDs(this.Count);
            target_shapeids.AddRange(shapeids);
            return target_shapeids;
        }

        public VisioAutomation.ShapeIDPairs ToShapeIDPairs()
        {
            if (this.Type == TargetShapesType.UseSelection)
            {
                throw new System.ArgumentException("This method only supported when the target shapes have been resolved");

            }

            if (this._shapes == null)
            {
                throw new System.ArgumentException("Target shapes must be resolved before calling ToShapeIDs()");
            }

            return VisioAutomation.ShapeIDPairs.FromShapes(this._shapes);
        }

        internal int SelectShapesAndCount(VisioScripting.Client client)
        {
            client.Application.AssertHasActiveApplication();

            var app = client.Application.GetActiveApplication();
            var active_window = app.ActiveWindow;
            var sel = active_window.Selection;

            if (this._shapes == null)
            {
                int n = sel.Count;
                client.Output.WriteVerbose("GetTargetSelectionCount: Using active selection of {0} shapes", n);
                return n;
            }

            client.Output.WriteVerbose("GetTargetSelectionCount: Resetting selection to specified {0} shapes", this._shapes.Count);

            // Force empty slection
            active_window.DeselectAll();
            active_window.DeselectAll(); // doing this twice is deliberate

            // Force selection to specific shapes
            active_window.Select(this._shapes, IVisio.VisSelectArgs.visSelect);

            int selected_count = sel.Count;
            return selected_count;
        }

        private IList<IVisio.Shape> _resolve_shapes(VisioScripting.Client client)
        {
            client.Application.AssertHasActiveApplication();

            if (this._shapes == null)
            {
                var out_shapes = client.Selection.GetShapesInSelection();
                client.Output.WriteVerbose("GetTargetShapes: Returning {0} shapes from the active selection", out_shapes.Count);
                return out_shapes;
            }

            client.Output.WriteVerbose("GetTargetShapes: Returning {0} shapes that were passed in", this.Count);
            return this._shapes;
        }

        public TargetShapes  ResolveShapes(VisioScripting.Client client)
        {
            var shapes = this._resolve_shapes(client);
            var targetshapes = new TargetShapes(shapes);
            return targetshapes;
        }

        internal TargetShapes ResolveShapes2D(VisioScripting.Client client)
        {
            var shapes = this._resolve_shapes(client);
            var shapes_2d = shapes.Where(s => s.OneD == 0).ToList();
            var targetshapes = new TargetShapes(shapes_2d);
            return targetshapes;
        }
    }
}