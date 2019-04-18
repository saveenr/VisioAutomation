using VisioAutomation.Exceptions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting
{
    public class CommandTarget
    {
        private readonly Client _client;
        public IVisio.Application Application { get; private set; }
        public IVisio.Document ActiveDocument { get; private set; }
        public IVisio.Page ActivePage { get; private set; }

        public CommandTarget(Client client, CommandTargetFlags flags)
        {
            this._client = client;

            _check(flags);
        }

        private void _check(CommandTargetFlags flags)
        {
            bool require_app = (flags & CommandTargetFlags.RequireApplication) != 0;
            bool require_document = (flags & CommandTargetFlags.RequireDocument) != 0;
            bool require_page = (flags & CommandTargetFlags.RequirePage) != 0;

            require_app = require_app || require_document || require_page;
            require_document = require_document || require_page;

            this.Application = this._client.Application.GetAttachedApplication();

            if (require_app && this.Application == null)
            {
                string msg = string.Format("{0}: No Visio Application available", nameof(CommandTarget));
                throw new System.ArgumentException(msg);
            }

            if (require_document && this.ActiveDocument == null)
            {
                var doc = this.Application.ActiveDocument;

                string errmsg;
                this.ActiveDocument = doc;

                bool is_drawing = IsDocumentADrawing(this.Application, this.ActiveDocument, out errmsg);

                if (is_drawing)
                {
                    this._client.Output.WriteVerbose("{0}: Verified a drawing is available for use",nameof(CommandTarget));
                }
                else
                {
                    string msg = string.Format("{0}: The Document is not a drawing document", nameof(CommandTarget));
                    throw new VisioOperationException(msg);
                }
            }

            if (require_document && this.ActiveDocument == null)
            {
                string msg = string.Format("{0}: No Document", nameof(CommandTarget));
                throw new VisioOperationException(msg);
            }

            if (require_page && this.ActivePage == null )
            {
                if (this.Application == null)
                {
                    string msg = string.Format("{0}: Internal error application should never be null in this case", nameof(CommandTarget));
                    throw new VisioOperationException(msg);
                }
                this.ActivePage = this.Application.ActivePage;
            }

            if (require_page && this.ActivePage == null)
            {
                string msg = string.Format("{0}: No Page", nameof(CommandTarget));
                throw new VisioOperationException(msg);
            }
        }

        public static bool IsDocumentADrawing(IVisio.Application app, IVisio.Document doc, out string errmsg)
        {
            if (app == null)
            {
                throw new System.ArgumentNullException(nameof(app));
            }

            if (doc == null)
            {
                throw new System.ArgumentNullException(nameof(doc));
            }

            if (doc.Application != app)
            {
                string msg = string.Format("{0}: application does not match doc.application", nameof(CommandTarget));
                throw new System.ArgumentException(msg);
            }


            var active_window = app.ActiveWindow;

            // If there's no active window there can't be an active document
            if (active_window == null)
            {
                errmsg = string.Format("{0}: No Active Window", nameof(CommandTarget));
                return false;
            }

            // Check if the window type matches that of a document
            short active_window_type = active_window.Type;
            var vis_drawing = (int) IVisio.VisWinTypes.visDrawing;
            var vis_master = (int) IVisio.VisWinTypes.visMasterWin;
            // var vis_sheet = (short)IVisio.VisWinTypes.visSheet;

            // this.Client.Output.WriteVerbose("ActiveDocumentIsDrawing: The Active Window: Type={0} & SybType={1}", active_window_type, active_window.SubType);
            if (!(active_window_type == vis_drawing || active_window_type == vis_master))
            {
                errmsg = string.Format("{0}: The Active Window Type must be one of {1} or {2}", nameof(CommandTarget), IVisio.VisWinTypes.visDrawing, IVisio.VisWinTypes.visMasterWin);
                return false;
            }

            var ap = app.ActivePage;
            //  verify there is an active page
            if (ap == null)
            {
                // 64 means master is being edited
                if (active_window.SubType != 64)
                {
                    errmsg = string.Format("{0}: Window is not editing a master", nameof(CommandTarget));
                    return false;
                }
            }

            errmsg = null;
            return true;
        }
    }
}
