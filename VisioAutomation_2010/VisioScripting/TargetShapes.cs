using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioScripting
{
    public class TargetShapes : TargetObjects<IVisio.Shape>
    {
        
        private TargetShapes() : base()
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

        public VisioAutomation.Core.ShapeIDPairs ToShapeIDPairs()
        {
            var items = this._get_items_safe();

            return VisioAutomation.Core.ShapeIDPairs.FromShapes(items);
        }

        public TargetShapes ResolveToShapes(VisioScripting.Client client)
        {
            if (this.Resolved)
            {
                return this;
            }

            var cmdtarget = client.GetCommandTarget(CommandTargetFlags.RequireDocument); var active_window = cmdtarget.Application.ActiveWindow;
            var selection = active_window.Selection;
            var shapes = selection.ToList();
            var targetshapes = new TargetShapes(shapes);

            client.Output.WriteVerbose("Resolving to selection (numshapes={0}) from active window (caption=\"{1}\")", shapes.Count, active_window.Caption);

            return targetshapes;
        }

        public void ResolveToSelection(VisioScripting.Client client)
        {
            // the purpose of this class is to handle those Visio operations that
            // don't explicitly take a list of shapes, but instead rely on the active selection

            var shapes = this._get_items_unsafe();

            if (shapes==null)
            {
                // do nothing - use the active selection
                return;
            }

            if (shapes.Count < 1)
            {
                throw new System.ArgumentOutOfRangeException("Shapes parameter must contain at least one shape");
            }

            client.Selection.SelectShapes(VisioScripting.TargetWindow.Auto, shapes);
        }

        public IList<IVisio.Shape> Shapes => this._get_items_safe();

        public static TargetShapes Auto = new TargetShapes();
    }
}