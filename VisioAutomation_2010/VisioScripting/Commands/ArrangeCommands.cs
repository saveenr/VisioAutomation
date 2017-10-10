using System.Collections.Generic;
using VisioAutomation.Exceptions;
using VisioAutomation.Shapes;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    [System.Flags]
    public enum CommandTargetFlags
    {
        Application,
        ActiveDocument,
        ActivePage
    }

    public class CommandTarget
    {
        public IVisio.Application Application;
        public IVisio.Document ActiveDocument;
        public IVisio.Page ActivePage;

        public bool has_app => this.Application != null;
        public bool has_doc => this.ActiveDocument != null;
        public bool has_page => this.ActivePage != null;
        public Client Client;

        public CommandTarget(Client client)
        {
            this.Client = client;
        }

        public CommandTarget(Client client, CommandTargetFlags flags )
        {
            this.Client = client;

            check(flags);
        }

        public void Assert(CommandTargetFlags flags)
        {
            check(flags);
        }

        private void check(CommandTargetFlags flags)
        {
            bool require_app = (flags & CommandTargetFlags.Application) != 0;
            bool require_document = (flags & CommandTargetFlags.ActiveDocument) != 0;
            bool require_page = (flags & CommandTargetFlags.ActivePage) != 0;

            require_app = require_app || require_document || require_page;
            require_document = require_document || require_page;

            this.Application = this.Client.Application.VisioApplication;

            if (this.Application == null && require_app )
            {
                var has_app = this.Client.Application.VisioApplication != null;
                if (!has_app)
                {
                    throw new System.ArgumentException("CommandTarget: No Visio Application available");
                }
            }

            if (require_app && this.Application == null)
            {
                throw new VisioOperationException("CommandTarget: No Application");
            }

            if ((this.ActiveDocument == null) && require_document)
            {
                var active_window = this.Application.ActiveWindow;

                // If there's no active window there can't be an active document
                if (active_window == null)
                {
                    this.Client.Output.WriteVerbose("CommandTarget: No Active Document");
                    throw new System.ArgumentException("CommandTarget: No Active Document");
                }

                // Check if the window type matches that of a document
                short active_window_type = active_window.Type;
                var vis_drawing = (int) IVisio.VisWinTypes.visDrawing;
                var vis_master = (int) IVisio.VisWinTypes.visMasterWin;
                // var vis_sheet = (short)IVisio.VisWinTypes.visSheet;

                this.Client.Output.WriteVerbose("CommandTarget: The Active Window: Type={0} & SybType={1}", active_window_type,
                    active_window.SubType);
                if (!(active_window_type == vis_drawing || active_window_type == vis_master))
                {
                    this.Client.Output.WriteVerbose("CommandTarget: The Active Window Type must be one of {0} or {1}",
                        IVisio.VisWinTypes.visDrawing, IVisio.VisWinTypes.visMasterWin);
                    throw new System.ArgumentException("CommandTarget: The Active Window Type must be one of {0} or {1}");
                }

                //  verify there is an active page

                if (this.Application.ActivePage == null)
                {
                    this.Client.Output.WriteVerbose("CommandTarget: Active Page is null");

                    if (active_window.SubType == 64)
                    {
                        // 64 means master is being edited
                    }
                    else
                    {
                        this.Client.Output.WriteVerbose("CommandTarget: Active Page is null");
                    }
                }

                this.Client.Output.WriteVerbose("CommandTarget: Verified a drawing is available for use");
                this.ActiveDocument = this.Application.ActiveDocument;
            }

            if (this.ActiveDocument == null && require_document)
            {
                throw new VisioOperationException("CommandTarget: No Document");
            }

            if ((this.ActivePage == null) && ((flags & CommandTargetFlags.ActivePage) != 0))
            {
                if (this.Application == null)
                {
                    throw new VisioOperationException("CommandTarget: Internal error application should never be null in this case");
                }
                this.ActivePage = this.Application.ActivePage;
            }

            if (this.ActivePage == null && require_page)
            {
                throw new VisioOperationException("CommandTarget: No Page");
            }

        }
    }

    public class ArrangeCommands : CommandSet
    {
        internal ArrangeCommands(Client client) :
            base(client)
        {

        }

        public void Nudge(VisioScripting.Models.TargetShapes targets, double dx, double dy)
        {
            if (dx == 0.0 && dy == 0.0)
            {
                return;
            }

            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);

            int shape_count = targets.SetSelectionGetSelectedCount(this._client);
            if (shape_count < 1)
            {
                return;
            }

            using (var undoscope = this._client.Application.NewUndoScope("Nudge"))
            {
                var window = cmdtarget.Application.ActiveWindow;
                var selection = window.Selection;
                var unitcode = Microsoft.Office.Interop.Visio.VisUnitCodes.visInches;

                // Move method: http://msdn.microsoft.com/en-us/library/ms367549.aspx   
                selection.Move(dx, dy, unitcode);
            }
        }

        private static void SendShapes(Microsoft.Office.Interop.Visio.Selection selection, Models.ShapeSendDirection dir)
        {

            if (dir == Models.ShapeSendDirection.ToBack)
            {
                selection.SendToBack();
            }
            else if (dir == Models.ShapeSendDirection.Backward)
            {
                selection.SendBackward();
            }
            else if (dir == Models.ShapeSendDirection.Forward)
            {
                selection.BringForward();
            }
            else if (dir == Models.ShapeSendDirection.ToFront)
            {
                selection.BringToFront();
            }
        }


        public void Send(VisioScripting.Models.TargetShapes targets, Models.ShapeSendDirection dir)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);

            int shape_count = targets.SetSelectionGetSelectedCount(this._client);
            if (shape_count < 1)
            {
                return;
            }

            var window = cmdtarget.Application.ActiveWindow;
            var selection = window.Selection;
            ArrangeCommands.SendShapes(selection, dir);
        }

        public void SetLock(VisioScripting.Models.TargetShapes targets, LockCells lockcells)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);

            targets = targets.ResolveShapes(this._client);
            if (targets.Shapes.Count < 1)
            {
                return;
            }

            var page = cmdtarget.ActivePage;
            var target_shapeids = targets.ToShapeIDs();
            var writer = new SidSrcWriter();

            foreach (int shapeid in target_shapeids.ShapeIDs)
            {
                lockcells.SetFormulas(writer, (short)shapeid);
            }

            using (var undoscope = this._client.Application.NewUndoScope("Set Lock Properties"))
            {
                writer.Commit(page);
            }
        }


        public Dictionary<int,LockCells> GetLock(VisioScripting.Models.TargetShapes targets)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);

            targets = targets.ResolveShapes(this._client);
            if (targets.Shapes.Count < 1)
            {
                return new Dictionary<int, LockCells>();
            }

            var dic = new Dictionary<int, LockCells>();

            var page = cmdtarget.ActivePage;
            var target_shapeids = targets.ToShapeIDs();

            var cells = VisioAutomation.Shapes.LockCells.GetCells(page, target_shapeids.ShapeIDs, CellValueType.Formula);

            for (int i = 0; i < target_shapeids.ShapeIDs.Count; i++)
            {
                var shapeid = target_shapeids.ShapeIDs[i];
                var cur_cells = cells[i];
                dic[shapeid] = cur_cells;
            }

            return dic;
        }

        public void SetSize(VisioScripting.Models.TargetShapes targets, double? w, double? h)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);

            targets = targets.ResolveShapes(this._client);
            if (targets.Shapes.Count < 1)
            {
                return;
            }

            var active_page = cmdtarget.ActivePage;
            var shapeids = targets.ToShapeIDs();
            var writer = new SidSrcWriter();
            foreach (int shapeid in shapeids.ShapeIDs)
            {
                if (w.HasValue && w.Value>=0)
                {
                    writer.SetFormula((short)shapeid, VisioAutomation.ShapeSheet.SrcConstants.XFormWidth, w.Value);
                }
                if (h.HasValue && h.Value >= 0)
                {
                    writer.SetFormula((short)shapeid, VisioAutomation.ShapeSheet.SrcConstants.XFormHeight, h.Value);                    
                }
            }

            using (var undoscope = this._client.Application.NewUndoScope("Set Shape Size"))
            {
                writer.Commit(active_page);
            }
        }
    }
}