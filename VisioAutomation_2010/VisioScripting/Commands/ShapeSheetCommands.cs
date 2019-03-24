using System.Collections.Generic;
using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class ShapeSheetCommands : CommandSet
    {
        internal ShapeSheetCommands(Client client) :
            base(client)
        {

        }

        internal void __SetCells(Models.TargetShapes targets, VisioAutomation.ShapeSheet.CellGroups.CellGroup cells, IVisio.Page page)
        {
            targets = targets.ResolveShapes(this._client);
            var shape_ids = targets.ToShapeIDs();
            var writer = new SidSrcWriter();

            foreach (var shape_id in shape_ids.ShapeIDs)
            {
                if (cells is VisioAutomation.ShapeSheet.CellGroups.CellGroup)
                {
                    var cells_mr = (VisioAutomation.ShapeSheet.CellGroups.CellGroup)cells;
                    writer.SetValues((short)shape_id, cells_mr, 0);
                }
                else
                {
                    var cells_sr = (VisioAutomation.ShapeSheet.CellGroups.CellGroup)cells;
                    writer.SetValues((short)shape_id, cells_sr);

                }
            }

            writer.CommitFormulas(page);
        }

        public void SetShapeName(Models.TargetShapes targets, IList<string> names)
        {
            if (names == null || names.Count < 1)
            {
                // do nothing
                return;
            }

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return;
            }

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(SetShapeName)))
            {
                int numnames = names.Count;

                int up_to = System.Math.Min(numnames, targets.Shapes.Count);

                for (int i = 0; i < up_to; i++)
                {
                    var new_name = names[i];

                    if (new_name != null)
                    {
                        var shape = targets.Shapes[i];
                        shape.Name = new_name;
                    }
                }
            }
        }
        
        public VisioAutomation.SurfaceTarget GetShapeSheetSurface()
        {
            var drawing_surface = this._client.Draw.GetActiveDrawingSurface();
            var shapesheet_surface = drawing_surface;
            return shapesheet_surface;
        }
        
        public Models.ShapeSheetWriter GetWriterForPage(IVisio.Page page)
        {
            var writer = new Models.ShapeSheetWriter(this._client, page);
            return writer;
        }

        public Models.ShapeSheetReader GetReaderForPage(IVisio.Page page)
        {
            var reader = new Models.ShapeSheetReader(this._client, page);
            return reader;
        }
    }
}