using System.Collections.Generic;
using System.Linq;
using VisioScripting.Commands;
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
        public Commands.ChartingCommands Charting { get; }
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
        public Commands.LockCommands Lock { get; }
        public Commands.MasterCommands Master { get; }
        public Commands.OutputCommands Output { get; }
        public Commands.PageCommands Page { get; }
        public Commands.SelectionCommands Selection { get; }
        public Commands.ShapeSheetCommands ShapeSheet { get; }
        public Commands.TextCommands Text { get; }
        public Commands.UndoCommands Undo { get; }
        public Commands.UserDefinedCellCommands UserDefinedCell { get; }
        public Commands.ViewCommands View { get; }

        public Client(IVisio.Application app):
            this(app,new DefaultClientContext())
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
            this.Charting = new Commands.ChartingCommands(this);
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
            this.Lock = new Commands.LockCommands(this);
            this.Master = new Commands.MasterCommands(this);
            this.Output = new Commands.OutputCommands(this);
            this.Page = new Commands.PageCommands(this);
            this.Selection = new Commands.SelectionCommands(this);
            this.ShapeSheet = new Commands.ShapeSheetCommands(this);
            this.Text = new Commands.TextCommands(this);
            this.Undo = new Commands.UndoCommands(this);
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

        public ClientContext ClientContext
        {
            get { return this._client_context; }
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

        public CommandTarget GetCommandTargetPage()
        {
            var flags = CommandTargetRequirementFlags.RequireApplication | 
                        CommandTargetRequirementFlags.RequireActiveDocument |
                        CommandTargetRequirementFlags.RequirePage;
            var command_target = new CommandTarget(this, flags);
            return command_target;
        }

        public CommandTarget GetCommandTargetDocument()
        {
            var flags = CommandTargetRequirementFlags.RequireApplication | 
                        CommandTargetRequirementFlags.RequireActiveDocument;
            var command_target = new CommandTarget(this, flags);
            return command_target;
        }

        public CommandTarget GetCommandTargetApplication()
        {
            var flags = CommandTargetRequirementFlags.RequireApplication;
            var command_target = new CommandTarget(this, flags);
            return command_target;
        }

        private static List<string> _static_dlls;
        public List<string> Assemblies
        {
            get
            {
                if (_static_dlls==null)
                {
                    _static_dlls = new List<string>();
                    var type = typeof(VisioScripting.Client);
                    string path = System.IO.Path.GetDirectoryName(type.Assembly.Location);
                    _static_dlls.Add(System.IO.Path.Combine(path, "VisioAutomation.dll"));
                    _static_dlls.Add(System.IO.Path.Combine(path, "VisioAutomation.Models.dll"));
                    // dlls.Add(System.IO.Path.Combine(path, "VisioPS.dll"));
                    _static_dlls.Add(System.IO.Path.Combine(path, "VisioScripting.dll"));
                    _static_dlls.Add(System.IO.Path.Combine(path, "Microsoft.Msagl.dll"));
                    _static_dlls.Add(System.IO.Path.Combine(path, "GenTreeOps.dll"));
                }
                return _static_dlls;
            }
        }
    }
}