using System.Linq;
using VisioAutomation.Extensions;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "Shape")]
    public class Get_Shape : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        public GetShapeFlags Flags = GetShapeFlags.Selected;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            if (this.Flags == GetShapeFlags.Selected)
            {
                var shapes = scriptingsession.Selection.GetShapes(VisioAutomation.ShapesEnumeration.Flat);
                this.WriteObject(shapes);
            }
            else if (this.Flags == GetShapeFlags.SelectedNested)
            {
                var shapes = scriptingsession.Selection.GetShapes(VisioAutomation.ShapesEnumeration.ExpandGroups);
                this.WriteObject(shapes);
            }
            else if (this.Flags == GetShapeFlags.Page)
            {
                var application = scriptingsession.VisioApplication;
                var active_page = application.ActivePage;
                var shapes1 = active_page.Shapes;
                var shapes = shapes1.AsEnumerable().ToList();
                this.WriteObject(shapes);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }
        }
    }
}