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
                this.WriteObject(shapes,true);
            }
            else if (this.Flags == GetVisioShapeFlags.SelectedNested)
            {
                var shapes = scriptingsession.Selection.GetShapesRecursive();
                this.WriteObject(shapes);
            }
            else if (this.Flags == GetVisioShapeFlags.Page)
            {
                var application = scriptingsession.VisioApplication;
                var active_page = application.ActivePage;
                var shapes1 = active_page.Shapes;
                var shapes = shapes1.AsEnumerable().ToList();
                this.WriteObject(shapes,true);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }
        }
    }
}