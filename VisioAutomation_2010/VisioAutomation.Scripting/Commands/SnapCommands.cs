using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Scripting.Layout;
using VisioAutomation.Scripting.Utilities;

namespace VisioAutomation.Scripting.Commands
{
    public class SnapCommands : CommandSet
    {
        internal SnapCommands(Client client) :
            base(client)
        {
        }

        public void SnapSize(TargetShapes targets, double w, double h)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = targets.ResolveShapes2DOnly(this._client);
            if (shapes.Shapes.Count < 1)
            {
                return;
            }

            var application = this._client.Application.Get();
            var page = application.ActivePage;
            var target_ids = targets.ToShapeIDs();
            using (var undoscope = this._client.Application.NewUndoScope("Snape Shape Sizes"))
            {
                var snapsize = new Drawing.Size(w, h);
                var minsize = new Drawing.Size(w, h);
                ArrangeHelper.SnapSize(page, target_ids, snapsize, minsize);
            }
        }

        public void SnapCorner(TargetShapes targets, double w, double h, SnapCornerPosition corner)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = targets.ResolveShapes2DOnly(this._client);

            if (shapes.Shapes.Count < 1)
            {
                return;
            }

            var application = this._client.Application.Get();
            var page = application.ActivePage;
            var target_ids = targets.ToShapeIDs();
            using (var undoscope = this._client.Application.NewUndoScope("Snap Shape Corner"))
            {
                ArrangeHelper.SnapCorner(page, target_ids, new Drawing.Size(w, h), corner);
            }
        }

        public void SnapSize(TargetShapes targets, Drawing.Size snapsize, Drawing.Size minsize)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = targets.ResolveShapes2DOnly(this._client);

            if (shapes.Shapes.Count < 1)
            {
                return;
            }

            var application = this._client.Application.Get();
            var page = application.ActivePage;
            var target_ids = targets.ToShapeIDs();
            using (var undoscope = this._client.Application.NewUndoScope("Snap Shape Size"))
            {
                ArrangeHelper.SnapSize(page, target_ids, snapsize, minsize);
            }
        }
    }
}