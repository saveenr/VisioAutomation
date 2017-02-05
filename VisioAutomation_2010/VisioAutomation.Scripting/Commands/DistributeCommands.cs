using VisioAutomation.Models;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Scripting.Utilities;

namespace VisioAutomation.Scripting.Commands
{
    public class DistributeCommands : CommandSet
    {
        internal DistributeCommands(Client client) :
            base(client)
        {

        }

        public void DistributeOnAxis(TargetShapes targets, Axis axis, double spacing)
        {
            if (!this._client.Document.HasActiveDocument)
            {
                return;
            }
            var page = this._client.Page.Get();
            targets = targets.ResolveShapes(this._client);
            var targetids = targets.ToShapeIDs();
            using (var undoscope = this._client.Application.NewUndoScope("Distribute on Axis"))
            {
                ArrangeHelper.DistributeWithSpacing(page, targetids, axis, spacing);
            }
        }

        public void DistributeOnAxis(TargetShapes targets, Axis axis)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            int shape_count = targets.SetSelectionGetSelectedCount(this._client);
            if (shape_count < 1)
            {
                return;
            }


            IVisio.VisUICmds cmd;

            switch (axis)
            {
                case Axis.XAxis:
                    cmd = IVisio.VisUICmds.visCmdDistributeHSpace;
                    break;
                case Axis.YAxis:
                    cmd = IVisio.VisUICmds.visCmdDistributeVSpace;
                    break;
                default:
                    throw new System.ArgumentOutOfRangeException();
            }

            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Distribute on Axis"))
            {
                application.DoCmd((short)cmd);
            }
        }

        public void DistributeHorizontal(TargetShapes targets, AlignmentHorizontal halign)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            int shape_count = targets.SetSelectionGetSelectedCount(this._client);
            if (shape_count < 1)
            {
                return;
            }

            IVisio.VisUICmds cmd;

            switch (halign)
            {
                case AlignmentHorizontal.Left:
                    cmd = IVisio.VisUICmds.visCmdDistributeLeft;
                    break;
                case AlignmentHorizontal.Center:
                    cmd = IVisio.VisUICmds.visCmdDistributeCenter;
                    break;
                case AlignmentHorizontal.Right:
                    cmd = IVisio.VisUICmds.visCmdDistributeRight;
                    break;
                default: throw new System.ArgumentOutOfRangeException();
            }

            var application = this._client.Application.Get();
            application.DoCmd((short)cmd);
        }

        public void DistributeVertical(TargetShapes targets, AlignmentVertical valign)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            int shape_count = targets.SetSelectionGetSelectedCount(this._client);
            if (shape_count < 1)
            {
                return;
            }

            IVisio.VisUICmds cmd;
            switch (valign)
            {
                case AlignmentVertical.Top:
                    cmd = IVisio.VisUICmds.visCmdDistributeTop;
                    break;
                case AlignmentVertical.Center: cmd = IVisio.VisUICmds.visCmdDistributeMiddle; break;
                case AlignmentVertical.Bottom: cmd = IVisio.VisUICmds.visCmdDistributeBottom; break;
                default: throw new System.ArgumentOutOfRangeException();
            }

            var application = this._client.Application.Get();
            application.DoCmd((short)cmd);
        }
    }
}