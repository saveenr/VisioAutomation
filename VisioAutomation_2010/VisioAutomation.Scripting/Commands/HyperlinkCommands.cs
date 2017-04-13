using System.Collections.Generic;
using VisioAutomation.Scripting.Models;
using VisioAutomation.Shapes;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.Commands
{
    public class HyperlinkCommands : CommandSet
    {
        internal HyperlinkCommands(Client client) :
            base(client)
        {

        }

        public List<int> Add(TargetShapes targets, HyperlinkCells ctrl)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

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

            using (var undoscope = this._client.Application.NewUndoScope("Add Control"))
            {
                foreach (var shape in targets.Shapes)
                {
                    int hi = HyperlinkHelper.Add(shape, ctrl);
                    hyperlink_indices.Add(hi);
                }
            }

            return hyperlink_indices;
        }

        public void Delete(TargetShapes targets, int n)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return;
            }

            using (var undoscope = this._client.Application.NewUndoScope("Delete Control"))
            {
                foreach (var shape in targets.Shapes)
                {
                    HyperlinkHelper.Delete(shape, n);
                }
            }
        }

        public Dictionary<IVisio.Shape, IList<HyperlinkCells>> Get(TargetShapes targets)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return new Dictionary<IVisio.Shape, IList<HyperlinkCells>>(0);
            }

            var dic = new Dictionary<IVisio.Shape, IList<HyperlinkCells>>();
            foreach (var shape in targets.Shapes)
            {
                var hyperlinks = HyperlinkCells.GetCells(shape);
                dic[shape] = hyperlinks;
            }
            return dic;
        }
    }
}