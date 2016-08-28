using VisioAutomation.Drawing.Layout;

namespace VisioAutomation.Scripting.Commands
{
    public class DistributeCommands : CommandSet
    {
        internal DistributeCommands(Client client) :
            base(client)
        {

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

            Microsoft.Office.Interop.Visio.VisUICmds cmd;

            switch (halign)
            {
                case AlignmentHorizontal.Left:
                    cmd = Microsoft.Office.Interop.Visio.VisUICmds.visCmdDistributeLeft;
                    break;
                case AlignmentHorizontal.Center:
                    cmd = Microsoft.Office.Interop.Visio.VisUICmds.visCmdDistributeCenter;
                    break;
                case AlignmentHorizontal.Right:
                    cmd = Microsoft.Office.Interop.Visio.VisUICmds.visCmdDistributeRight;
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

            Microsoft.Office.Interop.Visio.VisUICmds cmd;
            switch (valign)
            {
                case AlignmentVertical.Top:
                    cmd = Microsoft.Office.Interop.Visio.VisUICmds.visCmdDistributeTop;
                    break;
                case AlignmentVertical.Center: cmd = Microsoft.Office.Interop.Visio.VisUICmds.visCmdDistributeMiddle; break;
                case AlignmentVertical.Bottom: cmd = Microsoft.Office.Interop.Visio.VisUICmds.visCmdDistributeBottom; break;
                default: throw new System.ArgumentOutOfRangeException();
            }

            var application = this._client.Application.Get();
            application.DoCmd((short)cmd);
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


            Microsoft.Office.Interop.Visio.VisUICmds cmd;

            switch (axis)
            {
                case Axis.XAxis:
                    cmd = Microsoft.Office.Interop.Visio.VisUICmds.visCmdDistributeHSpace;
                    break;
                case Axis.YAxis:
                    cmd = Microsoft.Office.Interop.Visio.VisUICmds.visCmdDistributeVSpace;
                    break;
                default:
                    throw new System.ArgumentOutOfRangeException();
            }

            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Distribute Shapes"))
            {
                application.DoCmd((short)cmd);
            }
        }
    }
}