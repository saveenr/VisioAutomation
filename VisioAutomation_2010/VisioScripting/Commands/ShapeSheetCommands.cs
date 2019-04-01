using System.Collections.Generic;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class ShapeSheetCommands : CommandSet
    {
        internal ShapeSheetCommands(Client client) :
            base(client)
        {

        }

        internal void __SetCells(Models.TargetShapes targetshapes, VASS.CellGroups.CellGroup cellgroup, IVisio.Page page)
        {
            targetshapes = targetshapes.ResolveShapes(this._client);
            var targetshapeids = targetshapes.ToShapeIDs();
            var writer = new VASS.Writers.SidSrcWriter();

            foreach (var shapeid in targetshapeids)
            {
                var cells_mr = (VASS.CellGroups.CellGroup)cellgroup;
                writer.SetValues((short)shapeid, cells_mr, 0);
            }

            writer.Commit(page, VASS.CellValueType.Formula);
        }

        public void SetShapeName(Models.TargetShapes targetshapes, IList<string> names)
        {
            if (names == null || names.Count < 1)
            {
                // do nothing
                return;
            }

            targetshapes = targetshapes.ResolveShapes(this._client);

            if (targetshapes.Count < 1)
            {
                return;
            }

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(SetShapeName)))
            {
                int numnames = names.Count;

                int up_to = System.Math.Min(numnames, targetshapes.Count);

                for (int i = 0; i < up_to; i++)
                {
                    var new_name = names[i];

                    if (new_name != null)
                    {
                        var shape = targetshapes[i];
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