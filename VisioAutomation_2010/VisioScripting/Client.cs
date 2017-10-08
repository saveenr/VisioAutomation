using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting
{
    public class Client
    {
        private readonly ClientContext _client_context;

        public Commands.AlignCommands Align { get; }
        public Commands.ApplicationCommands Application { get; }
        public Commands.ApplicationWindowCommands Window { get; }
        public Commands.ArrangeCommands Arrange { get; }
        public Commands.ConnectionCommands Connection { get; }
        public Commands.ConnectionPointCommands ConnectionPoint { get; }
        public Commands.ControlCommands Control { get; }
        public Commands.CustomPropertyCommands CustomProperty { get; }
        public Commands.DeveloperCommands Developer { get; }
        public Commands.DistributeCommands Distribute { get; }
        public Commands.DocumentCommands Document { get; }
        public Commands.DrawCommands Draw { get; }
        public Commands.ExportPageCommands ExportPage { get; }
        public Commands.ExportSelectionCommands ExportSelection { get; }
        public Commands.GroupingCommands Grouping { get; }
        public Commands.HyperlinkCommands Hyperlink { get; }
        public Commands.LayerCommands Layer { get; }
        public Commands.MasterCommands Master { get; }
        public Commands.OutputCommands Output { get; }
        public Commands.PageCommands Page { get; }
        public Commands.SelectionCommands Selection { get; }
        public Commands.ShapeSheetCommands ShapeSheet { get; }
        public Commands.TextCommands Text { get; }
        public Commands.UserDefinedCellCommands UserDefinedCell { get; }
        public Commands.ViewCommands View { get; }

        public bool VerboseLogging = true;

        public Client(IVisio.Application app):
            this(app,new VisioScripting.Models.DefaultClientContext())
        {
        }
        
        public Client(IVisio.Application app, ClientContext client_context)
        {
            if (client_context == null)
            {
                throw new System.ArgumentNullException(nameof(client_context));
            }

            this._client_context = client_context;

            this.Align = new Commands.AlignCommands(this);
            this.Application = new Commands.ApplicationCommands(this, app);
            this.Arrange = new Commands.ArrangeCommands(this);
            this.Connection = new Commands.ConnectionCommands(this);
            this.ConnectionPoint = new Commands.ConnectionPointCommands(this);
            this.Control = new Commands.ControlCommands(this);
            this.CustomProperty = new Commands.CustomPropertyCommands(this);
            this.Developer = new Commands.DeveloperCommands(this);
            this.Distribute = new Commands.DistributeCommands(this);
            this.Document = new Commands.DocumentCommands(this);
            this.Draw = new Commands.DrawCommands(this);
            this.ExportPage = new Commands.ExportPageCommands(this);
            this.ExportSelection = new Commands.ExportSelectionCommands(this);
            this.Grouping = new Commands.GroupingCommands(this);
            this.Hyperlink = new Commands.HyperlinkCommands(this);
            this.Layer = new Commands.LayerCommands(this);
            this.Master = new Commands.MasterCommands(this);
            this.Output = new Commands.OutputCommands(this);
            this.Page = new Commands.PageCommands(this);
            this.Selection = new Commands.SelectionCommands(this);
            this.ShapeSheet = new Commands.ShapeSheetCommands(this);
            this.Text = new Commands.TextCommands(this);
            this.UserDefinedCell = new Commands.UserDefinedCellCommands(this);
            this.View = new Commands.ViewCommands(this);
            this.Window = new Commands.ApplicationWindowCommands(this);
        }

        public System.Reflection.Assembly GetVisioAutomationAssembly()
        {
            var type = typeof(VisioAutomation.ShapeSheet.Src);
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
            this._client_context.WriteUser(s);
        }

        public void WriteDebug(string fmt, params object[] items)
        {
            string s = string.Format(fmt, items);
            this._client_context.WriteDebug(s);
        }

        public void WriteVerbose(string fmt, params object[] items)
        {
            if (this.VerboseLogging)
            {
                string s = string.Format(fmt, items);
                this._client_context.WriteVerbose(s);
            }
        }

        public void WriteWarning(string fmt, params object[] items)
        {
            string s = string.Format(fmt, items);
            this._client_context.WriteWarning(s);
        }

        public void WriteError(string fmt, params object[] items)
        {
            string s = string.Format(fmt, items);
            this._client_context.WriteError(s);
        }

        public void WriteUser(string s)
        {
            this._client_context.WriteUser(s);
        }

        public void WriteDebug(string s)
        {
            this._client_context.WriteDebug(s);
        }

        public void WriteVerbose(string s)
        {
            if (this.VerboseLogging)
            {
                this._client_context.WriteVerbose(s);
            }
        }

        public void WriteWarning(string s)
        {
            this._client_context.WriteWarning(s);
        }
        
        public void WriteError(string s)
        {
            this._client_context.WriteError(s);
        }
        
        internal static List<System.Reflection.PropertyInfo> GetProperties()
        {
            var commandset_t = typeof (Commands.CommandSet);
            var all_props = typeof(Client).GetProperties();
            var command_props = all_props
                .Where(p => commandset_t.IsAssignableFrom(p.PropertyType))
                .ToList();
            return command_props;
        }
    }
}