using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting
{
    public class Client
    {

        private Context _context;

        public Commands.ApplicationCommands Application { get; private set; }
        public Commands.ViewCommands View { get; private set; }
        public Commands.FormatCommands Format { get; private set; }
        public Commands.LayerCommands Layer { get; private set; }
        public Commands.ControlCommands Control { get; private set; }
        public Commands.CustomPropCommands CustomProp { get; private set; }
        public Commands.ExportCommands Export { get; private set; }
        public Commands.ConnectionCommands Connection { get; private set; }
        public Commands.ConnectionPointCommands ConnectionPoint { get; private set; }
        public Commands.DrawCommands Draw { get; private set; }
        public Commands.MasterCommands Master { get; private set; }
        public Commands.ArrangeCommands Arrange { get; private set; }
        public Commands.PageCommands Page { get; private set; }
        public Commands.SelectionCommands Selection { get; private set; }
        public Commands.ShapeSheetCommands ShapeSheet { get; private set; }
        public Commands.TextCommands Text { get; private set; }
        public Commands.UserDefinedCellCommands UserDefinedCell { get; private set; }
        public Commands.DocumentCommands Document { get; private set; }
        public Commands.DeveloperCommands Developer { get; private set; }
        public Commands.OutputCommands Output { get; private set; }

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

            this.Application = new Commands.ApplicationCommands(this,app);
            this.View = new Commands.ViewCommands(this);
            this.Format = new Commands.FormatCommands(this);
            this.Layer = new Commands.LayerCommands(this);
            this.Control = new Commands.ControlCommands(this);
            this.CustomProp = new Commands.CustomPropCommands(this);
            this.Export = new Commands.ExportCommands(this);
            this.Connection = new Commands.ConnectionCommands(this);
            this.ConnectionPoint = new Commands.ConnectionPointCommands(this);
            this.Draw = new Commands.DrawCommands(this);
            this.Master = new Commands.MasterCommands(this);
            this.Arrange = new Commands.ArrangeCommands(this);
            this.Page = new Commands.PageCommands(this);
            this.Selection = new Commands.SelectionCommands(this);
            this.ShapeSheet = new Commands.ShapeSheetCommands(this);
            this.Text = new Commands.TextCommands(this);
            this.UserDefinedCell = new Commands.UserDefinedCellCommands(this);
            this.Document = new Commands.DocumentCommands(this);
            this.Developer = new Commands.DeveloperCommands(this);
            this.Output = new Commands.OutputCommands(this);
        }

        public System.Reflection.Assembly GetVisioAutomationAssembly()
        {
            var type = typeof(ShapeSheet.SRC);
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
            get { return this._context; }
            set
            {
                if (value == null)
                {
                    string msg = "Context must be non-null";
                    throw new System.ArgumentException(msg);
                }
                this._context = value;
            }
        }

        internal static List<System.Reflection.PropertyInfo> GetProperties()
        {
            var commandset_t = typeof (CommandSet);
            var all_props = typeof(Client).GetProperties();
            var command_props = all_props
                .Where(p => commandset_t.IsAssignableFrom(p.PropertyType))
                .ToList();
            return command_props;
        }
    }
}