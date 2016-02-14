using System.Collections.Generic;
using VAHLINK = VisioAutomation.Shapes.Hyperlinks;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.Commands
{
    public class HyperlinkCommands : CommandSet
    {
        internal HyperlinkCommands(Client client) :
            base(client)
        {

        }

        public IList<int> Add(IList<IVisio.Shape> target_shapes, VAHLINK.HyperlinkCells ctrl)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            if (ctrl == null)
            {
                throw new System.ArgumentNullException(nameof(ctrl));
            }

            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return new List<int>(0);
            }

            var hyperlink_indices = new List<int>();

            using (var undoscope = this._client.Application.NewUndoScope("Add Control"))
            {
                foreach (var shape in shapes)
                {
                    int hi = VAHLINK.HyperlinkHelper.Add(shape, ctrl);
                    hyperlink_indices.Add(hi);
                }
            }

            return hyperlink_indices;
        }

        public void Delete(IList<IVisio.Shape> target_shapes, int n)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            }

            using (var undoscope = this._client.Application.NewUndoScope("Delete Control"))
            {
                foreach (var shape in shapes)
                {
                    VAHLINK.HyperlinkHelper.Delete(shape, n);
                }
            }
        }

        public Dictionary<IVisio.Shape, IList<VAHLINK.HyperlinkCells>> Get(IList<IVisio.Shape> target_shapes)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();
            
            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return new Dictionary<IVisio.Shape, IList<VAHLINK.HyperlinkCells>>(0);
            }

            var dic = new Dictionary<IVisio.Shape, IList<VAHLINK.HyperlinkCells>>();
            foreach (var shape in shapes)
            {
                var hyperlinks = VAHLINK.HyperlinkCells.GetCells(shape);
                dic[shape] = hyperlinks;
            }
            return dic;
        }
    }
}