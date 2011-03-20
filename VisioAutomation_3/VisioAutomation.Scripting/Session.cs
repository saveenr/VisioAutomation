using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Scripting
{
    public partial class Session
    {
        public IVisio.Application Application { get; set; }
        public SessionOptions Options { get; set; }

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
        public VA.Scripting.Commands.PageCommands Page{ get; private set; }
        public VA.Scripting.Commands.SelectionCommands Selection { get; private set; }
        public VA.Scripting.Commands.ShapeSheetCommands ShapeSheet{ get; private set; }
        public VA.Scripting.Commands.TextCommands Text { get; private set; }
        public VA.Scripting.Commands.UserDefinedCellCommands UserDefinedCell { get; private set; }
        public VA.Scripting.Commands.DocumentCommands Document { get; private set; }
        //public VA.Scripting.Commands.ControlCommands Control { get; private set; }
        //public VA.Scripting.Commands.ControlCommands Control { get; private set; }
        //public VA.Scripting.Commands.ControlCommands Control { get; private set; }
        //public VA.Scripting.Commands.ControlCommands Control { get; private set; }
        //public VA.Scripting.Commands.ControlCommands Control { get; private set; }
        //public VA.Scripting.Commands.ControlCommands Control { get; private set; }
        //public VA.Scripting.Commands.ControlCommands Control { get; private set; }
        //public VA.Scripting.Commands.ControlCommands Control { get; private set; }
        //public VA.Scripting.Commands.ControlCommands Control { get; private set; }
        //public VA.Scripting.Commands.ControlCommands Control { get; private set; }
        //public VA.Scripting.Commands.ControlCommands Control { get; private set; }
        //public VA.Scripting.Commands.ControlCommands Control { get; private set; }

        public Session() :
            this(null)
        {
        }

        public Session(IVisio.Application app)
        {
            this.Options = new SessionOptions();
            this.Application = app;
            this.View = new VA.Scripting.Commands.ViewCommands(this);
            this.Format = new Commands.FormatCommands(this);
            this.Layer = new Commands.LayerCommands(this);
            this.Control = new Commands.ControlCommands(this);
            this.CustomProp = new Commands.CustomPropCommands(this);
            this.Export = new Commands.ExportCommands(this);
            this.Connection = new Commands.ConnectionCommands(this);
            this.ConnectionPoint = new Commands.ConnectionPointCommands(this);
            this.Draw = new Commands.DrawCommands(this);
            this.Master= new Commands.MasterCommands(this);
            this.Layout = new Commands.LayoutCommands(this);
            this.Page = new Commands.PageCommands(this);
            this.Selection = new Commands.SelectionCommands(this);
            this.ShapeSheet = new Commands.ShapeSheetCommands(this);
            this.Text = new Commands.TextCommands(this);
            this.UserDefinedCell = new Commands.UserDefinedCellCommands(this);
            this.Document = new Commands.DocumentCommands(this);
            //this.Control = new Commands.ControlCommands(this);
            //this.Control = new Commands.ControlCommands(this);
            //this.Control = new Commands.ControlCommands(this);
            //this.Control = new Commands.ControlCommands(this);
            //this.Control = new Commands.ControlCommands(this);
            //this.Control = new Commands.ControlCommands(this);
            //this.Control = new Commands.ControlCommands(this);
            //this.Control = new Commands.ControlCommands(this);
            //this.Control = new Commands.ControlCommands(this);
            //this.Control = new Commands.ControlCommands(this);
            //this.Control = new Commands.ControlCommands(this);
            //this.Control = new Commands.ControlCommands(this);
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

        public System.Drawing.Size GetApplicationWindowSize()
        {
            var rect = Application.Window.GetWindowRect();
            var size = new System.Drawing.Size(rect.Width, rect.Height);
            return size;
        }

        /// <summary>
        /// Sets the width and height (in pixels) of the attached Visio application window
        /// </summary>
        /// <param name="scripting_session"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public void SetApplicationWindowSize(int width, int height)
        {
            if (width <= 0)
            {
                Write(OutputStream.Error, "width must be positive");
                return;
            }

            if (height <= 0)
            {
                Write(OutputStream.Error, "height must be positive");
                return;
            }

            var r = Application.Window.GetWindowRect();
            r.Width = width;
            r.Height = height;
            Application.Window.SetWindowRect(r);
        }

        public static IVisio.Application AttachToRunningApplication()
        {
            var app = ApplicationHelper.FindRunningApplication();
            if (app == null)
            {
                throw new AutomationException("Did not find a running instance of Visio 2007");
            }

            VA.UI.UserInterfaceHelper.BringApplicationWindowToFront(app);

            return app;
        }

        public void ForceApplicationClose()
        {
            var application = Application;
            var documents = application.Documents;
            VA.DocumentHelper.ForceCloseAll(documents);
            application.Quit(true);
            Application = null;
        }

        public void BringApplicationWindowToFront()
        {
            var app = Application;

            if (app == null)
            {
                return;
            }

            VA.UI.UserInterfaceHelper.BringApplicationWindowToFront(app);
        }

        public IVisio.Application StartNewApplication()
        {
            var app = new IVisio.ApplicationClass();
            Application = app;
            return app;
        }

        public bool HasSelectedShapes()
        {
            return HasSelectedShapes(1);
        }

        public bool HasSelectedShapes(int min_items)
        {
            Write(OutputStream.Verbose, "Checking for at least {0} selected shapes", min_items);
            if (min_items <= 0)
            {
                throw new System.ArgumentOutOfRangeException("min_items");
            }

            if (!HasActiveDrawing())
            {
                return false;
            }

            var application = Application;
            var active_window = application.ActiveWindow;
            var selection = active_window.Selection;
            bool v = selection.Count >= min_items;
            return v;
        }

        public bool HasActiveDrawing()
        {
            var application = Application;
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
            if (active_window.Type != (int)IVisio.VisWinTypes.visDrawing)
            {
                return false;
            }

            return true;
        }

        public void Undo()
        {
            Application.Undo();
        }

        public void Redo()
        {
            Application.Redo();
        }

        public void Duplicate()
        {
            if (!HasSelectedShapes())
            {
                return;
            }
            var active_window = this.View.GetActiveWindow();
            var selection = active_window.Selection;
            selection.Duplicate();
        }

        public string GetApplicationWindowText()
        {
            return VA.ApplicationHelper.GetApplicationWindowText(this.Application);
        }

    }
}