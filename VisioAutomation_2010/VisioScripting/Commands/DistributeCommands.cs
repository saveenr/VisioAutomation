using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class DistributeCommands : CommandSet
    {
        internal DistributeCommands(Client client) :
            base(client)
        {

        }

        public void DistributenOnAxis(TargetShapes targetshapes, Models.Axis axis, double spacing)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            var page = cmdtarget.ActivePage;
            targetshapes = targetshapes.Resolve(this._client);
            var targetshapeids = targetshapes.ToShapeIDs();
            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DistributeOnAxis)))
            {
                VisioScripting.Helpers.ArrangeHelper._distribute_with_spacing(page, targetshapeids, axis, spacing);
            }
        }

        public void DistributeOnAxis(VisioScripting.TargetSelection targetselection, Models.Axis axis)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

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

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DistributeOnAxis)))
            {
                cmdtarget.Application.DoCmd((short)cmd);
            }
        }

        public void DistributeHorizontal(TargetShapes targetshapes, Models.AlignmentHorizontal halign)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();

            int shape_count = targetshapes.SelectShapesAndCount(this._client);
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

        public void DistributeVertical(TargetShapes targetshapes, Models.AlignmentVertical valign)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();

            int shape_count = targetshapes.SelectShapesAndCount(this._client);
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