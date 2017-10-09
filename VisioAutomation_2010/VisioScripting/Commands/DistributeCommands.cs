using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class DistributeCommands : CommandSet
    {
        internal DistributeCommands(Client client) :
            base(client)
        {

        }

        public void DistributeOnAxis(VisioScripting.Models.TargetShapes targets, VisioScripting.Models.Axis axis, double spacing)
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
                VisioScripting.Helpers.ArrangeHelper.DistributeWithSpacing(page, targetids, axis, spacing);
            }
        }

        public void DistributeOnAxis(VisioScripting.Models.TargetShapes targets, VisioScripting.Models.Axis axis)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);

            int shape_count = targets.SetSelectionGetSelectedCount(this._client);
            if (shape_count < 1)
            {
                return;
            }


            IVisio.VisUICmds cmd;

            switch (axis)
            {
                case VisioScripting.Models.Axis.XAxis:
                    cmd = IVisio.VisUICmds.visCmdDistributeHSpace;
                    break;
                case VisioScripting.Models.Axis.YAxis:
                    cmd = IVisio.VisUICmds.visCmdDistributeVSpace;
                    break;
                default:
                    throw new System.ArgumentOutOfRangeException();
            }

            using (var undoscope = this._client.Application.NewUndoScope("Distribute on Axis"))
            {
                cmdtarget.Application.DoCmd((short)cmd);
            }
        }

        public void DistributeHorizontal(VisioScripting.Models.TargetShapes targets, VisioScripting.Models.AlignmentHorizontal halign)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);

            int shape_count = targets.SetSelectionGetSelectedCount(this._client);
            if (shape_count < 1)
            {
                return;
            }

            IVisio.VisUICmds cmd;

            switch (halign)
            {
                case VisioScripting.Models.AlignmentHorizontal.Left:
                    cmd = IVisio.VisUICmds.visCmdDistributeLeft;
                    break;
                case VisioScripting.Models.AlignmentHorizontal.Center:
                    cmd = IVisio.VisUICmds.visCmdDistributeCenter;
                    break;
                case VisioScripting.Models.AlignmentHorizontal.Right:
                    cmd = IVisio.VisUICmds.visCmdDistributeRight;
                    break;
                default: throw new System.ArgumentOutOfRangeException();
            }

            cmdtarget.Application.DoCmd((short)cmd);
        }

        public void DistributeVertical(VisioScripting.Models.TargetShapes targets, VisioScripting.Models.AlignmentVertical valign)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);

            int shape_count = targets.SetSelectionGetSelectedCount(this._client);
            if (shape_count < 1)
            {
                return;
            }

            IVisio.VisUICmds cmd;
            switch (valign)
            {
                case VisioScripting.Models.AlignmentVertical.Top:
                    cmd = IVisio.VisUICmds.visCmdDistributeTop;
                    break;
                case VisioScripting.Models.AlignmentVertical.Center: cmd = IVisio.VisUICmds.visCmdDistributeMiddle; break;
                case VisioScripting.Models.AlignmentVertical.Bottom: cmd = IVisio.VisUICmds.visCmdDistributeBottom; break;
                default: throw new System.ArgumentOutOfRangeException();
            }

            cmdtarget.Application.DoCmd((short)cmd);
        }
    }
}