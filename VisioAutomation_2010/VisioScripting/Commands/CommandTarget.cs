using VisioAutomation.Exceptions;

namespace VisioScripting.Commands
{
    public class CommandTarget
    {
        public readonly Client Client;
        public Microsoft.Office.Interop.Visio.Application Application;
        public Microsoft.Office.Interop.Visio.Document ActiveDocument;
        public Microsoft.Office.Interop.Visio.Page ActivePage;

        public bool HasActiveApplication => this.Application != null;
        public bool HasActiveDocument => this.ActiveDocument != null;
        public bool HasActivePage => this.ActivePage != null;


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
                var vis_drawing = (int) Microsoft.Office.Interop.Visio.VisWinTypes.visDrawing;
                var vis_master = (int) Microsoft.Office.Interop.Visio.VisWinTypes.visMasterWin;
                // var vis_sheet = (short)IVisio.VisWinTypes.visSheet;

                this.Client.Output.WriteVerbose("CommandTarget: The Active Window: Type={0} & SybType={1}", active_window_type,
                    active_window.SubType);
                if (!(active_window_type == vis_drawing || active_window_type == vis_master))
                {
                    this.Client.Output.WriteVerbose("CommandTarget: The Active Window Type must be one of {0} or {1}",
                        Microsoft.Office.Interop.Visio.VisWinTypes.visDrawing, Microsoft.Office.Interop.Visio.VisWinTypes.visMasterWin);
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
}