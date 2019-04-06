using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Shapes;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class HyperlinkCommands : CommandSet
    {
        internal HyperlinkCommands(Client client) :
            base(client)
        {

        }

        public List<int> AddHyperlink(TargetShapes targetshapes, HyperlinkCells hlink)
        {
            if (hlink == null)
            {
                throw new System.ArgumentNullException(nameof(hlink));
            }

            targetshapes = targetshapes.Resolve(this._client);
            
            if (targetshapes.Shapes.Count < 1)
            {
                return new List<int>(0);
            }

            var hyperlink_indices = new List<int>();

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(AddHyperlink)))
            {
                foreach (var shape in targetshapes.Shapes)
                {
                    int hi = HyperlinkHelper.Add(shape, hlink);
                    hyperlink_indices.Add(hi);
                }
            }

            return hyperlink_indices;
        }

        public void DeleteHyperlinkAtIndex(TargetShapes targetshapes, int n)
        {
            targetshapes = targetshapes.Resolve(this._client);

            if (targetshapes.Shapes.Count < 1)
            {
                return;
            }

            // restrict the operation to those shapes that actually have enough
            // controls to qualify for deleting 
            var qualified_shapes = targetshapes.Shapes.Where(shape => HyperlinkHelper.GetCount(shape) > n);

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DeleteHyperlinkAtIndex)))
            {
                foreach (var shape in qualified_shapes)
                {
                    HyperlinkHelper.Delete(shape, n);
                }
            }
        }

        public Dictionary<IVisio.Shape, IList<HyperlinkCells>> GetHyperlinks(TargetShapes targetshapes, VASS.CellValueType cvt)
        {
            targetshapes = targetshapes.Resolve(this._client);

            if (targetshapes.Shapes.Count < 1)
            {
                return new Dictionary<IVisio.Shape, IList<HyperlinkCells>>(0);
            }

            var dic = new Dictionary<IVisio.Shape, IList<HyperlinkCells>>();
            foreach (var shape in targetshapes.Shapes)
            {
                var hyperlinks = HyperlinkCells.GetCells(shape, cvt);
                dic[shape] = hyperlinks;
            }
            return dic;
        }
    }
}