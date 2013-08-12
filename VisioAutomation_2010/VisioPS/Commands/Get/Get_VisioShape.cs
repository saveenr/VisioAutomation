using System.Linq;
using VisioAutomation.Extensions;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioShape")]
    public class Get_VisioShape : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        public GetVisioShapeFlags Flags = GetVisioShapeFlags.Selected;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            if (this.Flags == GetVisioShapeFlags.Selected)
            {
                var shapes = scriptingsession.Selection.GetShapes();
                this.WriteObject(shapes,false);
            }
            else if (this.Flags == GetVisioShapeFlags.SelectedNested)
            {
                var shapes = scriptingsession.Selection.GetShapesRecursive();
                this.WriteObject(shapes, false);
            }
            else if (this.Flags == GetVisioShapeFlags.Page)
            {
                var application = scriptingsession.VisioApplication;
                var active_page = application.ActivePage;
                var page_shapes = active_page.Shapes;
                var shapes = page_shapes.AsEnumerable().ToList();
                this.WriteObject(shapes, false);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }
        }
    }
}