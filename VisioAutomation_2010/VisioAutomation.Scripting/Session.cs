using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using VisioAutomation.Scripting.Commands;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Scripting
{
    public class Session
    {
        public IVisio.Application VisioApplication { get; set; }
        public SessionOptions Options { get; set; }

        public ApplicationCommands Application { get; private set; }
        public ViewCommands View { get; private set; }
        public FormatCommands Format { get; private set; }
        public LayerCommands Layer { get; private set; }
        public ControlCommands Control { get; private set; }
        public CustomPropCommands CustomProp { get; private set; }
        public ExportCommands Export { get; private set; }
        public ConnectionCommands Connection { get; private set; }
        public ConnectionPointCommands ConnectionPoint { get; private set; }
        public DrawCommands Draw { get; private set; }
        public MasterCommands Master { get; private set; }
        public LayoutCommands Layout { get; private set; }
        public PageCommands Page { get; private set; }
        public SelectionCommands Selection { get; private set; }
        public ShapeSheetCommands ShapeSheet { get; private set; }
        public TextCommands Text { get; private set; }
        public UserDefinedCellCommands UserDefinedCell { get; private set; }
        public DocumentCommands Document { get; private set; }
        public DeveloperCommands Developer { get; private set; }
        public OutputCommands Output { get; private set; }

        public Session() :
            this(null)
        {
        }

        public Session(IVisio.Application app)
        {
            this.Options = new SessionOptions();
            this.VisioApplication = app;

            this.Application = new ApplicationCommands(this);
            this.View = new ViewCommands(this);
            this.Format = new FormatCommands(this);
            this.Layer = new LayerCommands(this);
            this.Control = new ControlCommands(this);
            this.CustomProp = new CustomPropCommands(this);
            this.Export = new ExportCommands(this);
            this.Connection = new ConnectionCommands(this);
            this.ConnectionPoint = new ConnectionPointCommands(this);
            this.Draw = new DrawCommands(this);
            this.Master = new MasterCommands(this);
            this.Layout = new LayoutCommands(this);
            this.Page = new PageCommands(this);
            this.Selection = new SelectionCommands(this);
            this.ShapeSheet = new ShapeSheetCommands(this);
            this.Text = new TextCommands(this);
            this.UserDefinedCell = new UserDefinedCellCommands(this);
            this.Document = new DocumentCommands(this);
            this.Developer = new DeveloperCommands(this);
            this.Output = new OutputCommands(this);
        }

        internal void Write(OutputStream output, string s)
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
                throw new ArgumentOutOfRangeException("output");
            }
        }

        internal void Write(OutputStream output, string fmt, params object[] items)
        {
            string s = String.Format(fmt, items);
            this.Write(output, s);
        }
        
        internal bool HasSelectedShapes()
        {
            return this.Selection.HasShapes();
        }

        internal bool HasSelectedShapes(int min_items)
        {
            return this.Selection.HasShapes(min_items);
        }

        public bool HasActiveDrawing
        {
            get
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

                if (active_window.Type != (int)IVisio.VisWinTypes.visDrawing)
                {
                    return false;
                }

                return true;
            }
        }

        internal static List<System.Reflection.PropertyInfo> GetCommandSetProperties()
        {
            var props = typeof(Scripting.Session).GetProperties()
                .Where(
                    p => typeof(Scripting.CommandSet).IsAssignableFrom(p.PropertyType))
                .ToList();
            return props;
        }
    }
}