using System.Collections.Generic;
using VisioAutomation.Shapes;
using VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class HyperlinkCommands : CommandSet
    {
        internal HyperlinkCommands(Client client) :
            base(client)
        {

        }

        public List<int> AddHyperlink(VisioScripting.Models.TargetShapes targets, HyperlinkCells ctrl)
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            if (ctrl == null)
            {
                throw new System.ArgumentNullException(nameof(ctrl));
            }

            targets = targets.ResolveShapes(this._client);
            
            if (targets.Shapes.Count < 1)
            {
                return new List<int>(0);
            }

            var hyperlink_indices = new List<int>();

            using (var undoscope = this._client.Application.NewUndoScope(nameof(AddHyperlink)))
            {
                foreach (var shape in targets.Shapes)
                {
                    int hi = HyperlinkHelper.Add(shape, ctrl);
                    hyperlink_indices.Add(hi);
                }
            }

            return hyperlink_indices;
        }

        public void DeleteHyperlinkAtIndex(VisioScripting.Models.TargetShapes targets, int n)
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return;
            }

            using (var undoscope = this._client.Application.NewUndoScope(nameof(DeleteHyperlinkAtIndex)))
            {
                foreach (var shape in targets.Shapes)
                {
                    HyperlinkHelper.Delete(shape, n);
                }
            }
        }

        public Dictionary<IVisio.Shape, IList<HyperlinkCells>> GetHyperlinkCells(VisioScripting.Models.TargetShapes targets, CellValueType cvt)
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return new Dictionary<IVisio.Shape, IList<HyperlinkCells>>(0);
            }

            var dic = new Dictionary<IVisio.Shape, IList<HyperlinkCells>>();
            foreach (var shape in targets.Shapes)
            {
                var hyperlinks = HyperlinkCells.GetCells(shape, cvt);
                dic[shape] = hyperlinks;
            }
            return dic;
        }
    }
}