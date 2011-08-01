using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Scripting
{
    public class Session
    {
        public IVisio.Application VisioApplication { get; set; }
        public SessionOptions Options { get; set; }

        public VA.Scripting.Commands.ApplicationCommands Application { get; private set; }
        public VA.Scripting.Commands.ViewCommands View { get; private set; }
        public VA.Scripting.Commands.FormatCommands Format { get; private set; }
        public VA.Scripting.Commands.LayerCommands Layer { get; private set; }
        public VA.Scripting.Commands.ControlCommands Control { get; private set; }
        public VA.Scripting.Commands.CustomPropCommands CustomProp { get; private set; }
        public VA.Scripting.Commands.ExportCommands Export { get; private set; }
        public VA.Scripting.Commands.ConnectionCommands Connection { get; private set; }
        public VA.Scripting.Commands.ConnectionPointCommands ConnectionPoint { get; private set; }
        public VA.Scripting.Commands.DrawCommands Draw { get; private set; }
        public VA.Scripting.Commands.MasterCommands Master { get; private set; }
        public VA.Scripting.Commands.LayoutCommands Layout { get; private set; }
        public VA.Scripting.Commands.PageCommands Page { get; private set; }
        public VA.Scripting.Commands.SelectionCommands Selection { get; private set; }
        public VA.Scripting.Commands.ShapeSheetCommands ShapeSheet { get; private set; }
        public VA.Scripting.Commands.TextCommands Text { get; private set; }
        public VA.Scripting.Commands.UserDefinedCellCommands UserDefinedCell { get; private set; }
        public VA.Scripting.Commands.DocumentCommands Document { get; private set; }
        public VA.Scripting.Commands.DeveloperCommands Developer { get; private set; }

        public Session() :
            this(null)
        {
        }

        public Session(IVisio.Application app)
        {
            this.Options = new SessionOptions();
            this.VisioApplication = app;

            this.Application = new Commands.ApplicationCommands(this);
            this.View = new VA.Scripting.Commands.ViewCommands(this);
            this.Format = new Commands.FormatCommands(this);
            this.Layer = new Commands.LayerCommands(this);
            this.Control = new Commands.ControlCommands(this);
            this.CustomProp = new Commands.CustomPropCommands(this);
            this.Export = new Commands.ExportCommands(this);
            this.Connection = new Commands.ConnectionCommands(this);
            this.ConnectionPoint = new Commands.ConnectionPointCommands(this);
            this.Draw = new Commands.DrawCommands(this);
            this.Master = new Commands.MasterCommands(this);
            this.Layout = new Commands.LayoutCommands(this);
            this.Page = new Commands.PageCommands(this);
            this.Selection = new Commands.SelectionCommands(this);
            this.ShapeSheet = new Commands.ShapeSheetCommands(this);
            this.Text = new Commands.TextCommands(this);
            this.UserDefinedCell = new Commands.UserDefinedCellCommands(this);
            this.Document = new Commands.DocumentCommands(this);
            this.Developer = new Commands.DeveloperCommands(this);
        }

        public void Write(OutputStream output, string s)
        {
            if (output == OutputStream.User)
            {
                this.Options.WriteUser(s);
            }
            else if (output == OutputStream.Error)
            {
                this.Options.WriteError(s);
            }
            else if (output == OutputStream.Debug)
            {
                this.Options.WriteDebug(s);
            }
            else if (output == OutputStream.Verbose)
            {
                this.Options.WriteVerbose(s);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException("output");
            }
        }

        public void Write(OutputStream output, string fmt, params object[] items)
        {
            string s = string.Format(fmt, items);
            this.Write(output, s);
        }


        internal bool HasSelectedShapes()
        {
            return this.Selection.HasSelectedShapes();
        }

        internal bool HasSelectedShapes(int min_items)
        {
            return this.Selection.HasSelectedShapes(min_items);
        }

        public bool HasActiveDrawing()
        {
            var application = VisioApplication;
            var active_window = application.ActiveWindow;

            if (active_window == null)
            {
                return false;
            }
            if (application.ActiveDocument == null)
            {
                return false;
            }
            if (application.ActivePage == null)
            {
                return false;
            }
            if (active_window.Type != (int) IVisio.VisWinTypes.visDrawing)
            {
                return false;
            }

            return true;
        }
    }
}