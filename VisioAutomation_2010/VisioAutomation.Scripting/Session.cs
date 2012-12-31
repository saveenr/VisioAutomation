using System;
using System.Collections.Generic;
using System.Linq;
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
        public VA.Scripting.Commands.OutputCommands Output { get; private set; }

        public Session() :
            this(null)
        {
        }

        public Session(IVisio.Application app)
        {
            this.Options = new SessionOptions();
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
            this.Layout = new VA.Scripting.Commands.LayoutCommands(this);
            this.Page = new VA.Scripting.Commands.PageCommands(this);
            this.Selection = new VA.Scripting.Commands.SelectionCommands(this);
            this.ShapeSheet = new VA.Scripting.Commands.ShapeSheetCommands(this);
            this.Text = new VA.Scripting.Commands.TextCommands(this);
            this.UserDefinedCell = new VA.Scripting.Commands.UserDefinedCellCommands(this);
            this.Document = new VA.Scripting.Commands.DocumentCommands(this);
            this.Developer = new VA.Scripting.Commands.DeveloperCommands(this);
            this.Output = new VA.Scripting.Commands.OutputCommands(this);
        }

        public void WriteUser(string fmt, params object[] items)
        {
            string s = String.Format(fmt, items);
            this.Options.WriteUser(s);
        }

        public void WriteDebug(string fmt, params object[] items)
        {
            string s = String.Format(fmt, items);
            this.Options.WriteDebug(s);
        }

        public void WriteVerbose(string fmt, params object[] items)
        {
            string s = String.Format(fmt, items);
            this.Options.WriteVerbose(s);
        }

        public void WriteError(string fmt, params object[] items)
        {
            string s = String.Format(fmt, items);
            this.Options.WriteError(s);
        }

        public void WriteUser(string s)
        {
            this.Options.WriteUser(s);
        }

        public void WriteDebug(string s)
        {
            this.Options.WriteDebug(s);
        }

        public void WriteVerbose(string s)
        {
            this.Options.WriteVerbose(s);
        }

        public void WriteError(string s)
        {
            this.Options.WriteError(s);
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