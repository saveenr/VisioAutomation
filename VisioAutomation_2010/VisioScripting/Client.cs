using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting
{
    public class Client
    {
        private readonly ClientContext _client_context;

        public Commands.ApplicationCommands Application { get; }
        public Commands.ArrangeCommands Arrange { get; }
        public Commands.ModelCommands Model { get; }
        public Commands.ConnectionCommands Connection { get; }
        public Commands.ConnectionPointCommands ConnectionPoint { get; }
        public Commands.ContainerCommands Container { get; }
        public Commands.ControlCommands Control { get; }
        public Commands.CustomPropertyCommands CustomProperty { get; }
        public Commands.DeveloperCommands Developer { get; }
        public Commands.DocumentCommands Document { get; }
        public Commands.DrawCommands Draw { get; }
        public Commands.ExportCommands Export { get; }
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
            this.Application = new Commands.ApplicationCommands(this, app);
            this.Arrange = new Commands.ArrangeCommands(this);
            this.Model = new Commands.ModelCommands(this);
            this.Connection = new Commands.ConnectionCommands(this);
            this.ConnectionPoint = new Commands.ConnectionPointCommands(this);
            this.Container = new Commands.ContainerCommands(this);
            this.Control = new Commands.ControlCommands(this);
            this.CustomProperty = new Commands.CustomPropertyCommands(this);
            this.Developer = new Commands.DeveloperCommands(this);
            this.Document = new Commands.DocumentCommands(this);
            this.Draw = new Commands.DrawCommands(this);
            this.Export = new Commands.ExportCommands(this);
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
            var flags = CommandTargetFlags.RequireApplication | 
                        CommandTargetFlags.RequireDocument |
                        CommandTargetFlags.RequirePage;
            var command_target = new CommandTarget(this, flags);
            return command_target;
        }

        public CommandTarget GetCommandTargetDocument()
        {
            var flags = CommandTargetFlags.RequireApplication | 
                        CommandTargetFlags.RequireDocument;
            var command_target = new CommandTarget(this, flags);
            return command_target;
        }

        public CommandTarget GetCommandTargetApplication()
        {
            var flags = CommandTargetFlags.RequireApplication;
            var command_target = new CommandTarget(this, flags);
            return command_target;
        }
    }
}