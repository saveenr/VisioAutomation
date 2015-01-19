using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Scripting
{
    public class Client
    {
        public IVisio.Application VisioApplication { get; set; }
        private Context _context;

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
        public VA.Scripting.Commands.ArrangeCommands Arrange { get; private set; }
        public VA.Scripting.Commands.PageCommands Page { get; private set; }
        public VA.Scripting.Commands.SelectionCommands Selection { get; private set; }
        public VA.Scripting.Commands.ShapeSheetCommands ShapeSheet { get; private set; }
        public VA.Scripting.Commands.TextCommands Text { get; private set; }
        public VA.Scripting.Commands.UserDefinedCellCommands UserDefinedCell { get; private set; }
        public VA.Scripting.Commands.DocumentCommands Document { get; private set; }
        public VA.Scripting.Commands.DeveloperCommands Developer { get; private set; }
        public VA.Scripting.Commands.OutputCommands Output { get; private set; }

        public Client(IVisio.Application app):
            this(app,new DefaultContext())
        {
        }
        
        public Client(IVisio.Application app, Context context)
        {
            if (context == null)
            {
                throw new System.ArgumentNullException();
            }
            this._context = context;
            this.VisioApplication = app;

            this.Application = new VA.Scripting.Commands.ApplicationCommands(this);
            this.View = new VA.Scripting.Commands.ViewCommands(this);
            this.Format = new VA.Scripting.Commands.FormatCommands(this);
            this.Layer = new VA.Scripting.Commands.LayerCommands(this);
            this.Control = new VA.Scripting.Commands.ControlCommands(this);
            this.CustomProp = new VA.Scripting.Commands.CustomPropCommands(this);
            this.Export = new VA.Scripting.Commands.ExportCommands(this);
            this.Connection = new VA.Scripting.Commands.ConnectionCommands(this);
            this.ConnectionPoint = new VA.Scripting.Commands.ConnectionPointCommands(this);
            this.Draw = new VA.Scripting.Commands.DrawCommands(this);
            this.Master = new VA.Scripting.Commands.MasterCommands(this);
            this.Arrange = new VA.Scripting.Commands.ArrangeCommands(this);
            this.Page = new VA.Scripting.Commands.PageCommands(this);
            this.Selection = new VA.Scripting.Commands.SelectionCommands(this);
            this.ShapeSheet = new VA.Scripting.Commands.ShapeSheetCommands(this);
            this.Text = new VA.Scripting.Commands.TextCommands(this);
            this.UserDefinedCell = new VA.Scripting.Commands.UserDefinedCellCommands(this);
            this.Document = new VA.Scripting.Commands.DocumentCommands(this);
            this.Developer = new VA.Scripting.Commands.DeveloperCommands(this);
            this.Output = new VA.Scripting.Commands.OutputCommands(this);
        }

        public System.Reflection.Assembly GetVisioAutomationAssembly()
        {
            var type = typeof(VA.ShapeSheet.SRC);
            var asm = type.Assembly;
            return asm;
        }

        public System.Reflection.Assembly GetVisioAssembly()
        {
            var type = typeof(IVisio.Application);
            var asm = type.Assembly;
            return asm;
        }
        
        public void WriteUser(string fmt, params object[] items)
        {
            string s = string.Format(fmt, items);
            this._context.WriteUser(s);
        }

        public void WriteDebug(string fmt, params object[] items)
        {
            string s = string.Format(fmt, items);
            this._context.WriteDebug(s);
        }

        public void WriteVerbose(string fmt, params object[] items)
        {
            string s = string.Format(fmt, items);
            this._context.WriteVerbose(s);
        }

        public void WriteWarning(string fmt, params object[] items)
        {
            string s = string.Format(fmt, items);
            this._context.WriteWarning(s);
        }

        public void WriteError(string fmt, params object[] items)
        {
            string s = string.Format(fmt, items);
            this._context.WriteError(s);
        }

        public void WriteUser(string s)
        {
            this._context.WriteUser(s);
        }

        public void WriteDebug(string s)
        {
            this._context.WriteDebug(s);
        }

        public void WriteVerbose(string s)
        {
            this._context.WriteVerbose(s);
        }

        public void WriteWarning(string s)
        {
            this._context.WriteWarning(s);
        }
        
        public void WriteError(string s)
        {
            this._context.WriteError(s);
        }

        public Context Context
        {
            get { return _context; }
            set
            {
                if (value == null)
                {
                    string msg = "Context must be non-null";
                    throw new System.ArgumentException(msg);
                }
                _context = value;
            }
        }

        internal static List<System.Reflection.PropertyInfo> GetCommandSetProperties()
        {
            var commandset_t = typeof (Scripting.CommandSet);
            var all_props = typeof(Scripting.Client).GetProperties();
            var command_props = all_props
                .Where(p => commandset_t.IsAssignableFrom(p.PropertyType))
                .ToList();
            return command_props;
        }
    }
}