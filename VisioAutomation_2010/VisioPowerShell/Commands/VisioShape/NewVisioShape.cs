using System.Collections.Generic;
using System.Linq;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioShape
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, Nouns.VisioShape)]
    public class NewVisioShape : VisioCmdlet
    {
        [SMA.Parameter(ParameterSetName = "masters", Position = 0, Mandatory = true)]
        public IVisio.Master[] Master { get; set; }

        [SMA.Parameter(ParameterSetName = "rectangle", Position = 0, Mandatory = true)]
        public SMA.SwitchParameter Rectangle { get; set; }

        [SMA.Parameter(ParameterSetName = "oval", Position = 0, Mandatory = true)]
        public SMA.SwitchParameter Oval { get; set; }

        [SMA.Parameter(ParameterSetName = "line", Position = 0, Mandatory = true)]
        public SMA.SwitchParameter Line { get; set; }

        [SMA.Parameter(ParameterSetName = "polyline", Position = 0, Mandatory = true)]
        public SMA.SwitchParameter Polyline { get; set; }

        [SMA.Parameter(ParameterSetName = "bezier", Position = 0, Mandatory = true)]
        public SMA.SwitchParameter Bezier { get; set; }


        [SMA.Parameter(ParameterSetName = "masters", Mandatory = true)]
        public VisioAutomation.Geometry.Point[] DropPosition { get; set; }

        [SMA.Parameter(ParameterSetName = "line", Mandatory = true)]
        public VisioAutomation.Geometry.Point From { get; set; }

        [SMA.Parameter(ParameterSetName = "line", Mandatory = true)]
        public VisioAutomation.Geometry.Point To { get; set; }


        [SMA.Parameter(ParameterSetName = "polyline", Mandatory = true)]
        [SMA.Parameter(ParameterSetName = "bezier", Mandatory = true)]
        public VisioAutomation.Geometry.Point[] Points { get; set; }

        [SMA.Parameter(ParameterSetName = "rectangle", Mandatory = true)]
        [SMA.Parameter(ParameterSetName = "oval", Mandatory = true)]
        public VisioAutomation.Geometry.Rectangle BoundingBox{ get; set; }

        [SMA.Parameter(Mandatory = false)]
        public VisioPowerShell.Models.ShapeCells[] Cells { get; set; }

        protected override void ProcessRecord()
        {
            if (this.Master != null)
            {
                _drop_shape();
            }
            else
            {
                _draw_shape();
            }
        }

        private void _draw_shape()
        {
            _check_num_Points();


            if (this.Rectangle)
            {
                var shape = this.Client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, this.BoundingBox);
                this.WriteObject(shape) ;
            }
            else if (this.Oval)
            {
                var shape = this.Client.Draw.DrawOval(VisioScripting.TargetPage.Auto, this.BoundingBox);
                this.WriteObject(shape);
            }
            else if (this.Line)
            {
                var shape = this.Client.Draw.DrawLine(VisioScripting.TargetPage.Auto, this.From, this.To);
                this.WriteObject(shape);
            }
            else if (this.Polyline)
            {
                var shape = this.Client.Draw.DrawPolyLine(VisioScripting.TargetPage.Auto, this.Points);
                this.WriteObject(shape);
            }
            else if (this.Bezier)
            {
                var shape = this.Client.Draw.DrawBezier(VisioScripting.TargetPage.Auto, this.Points);
                this.WriteObject(shape);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }
        }

        private void _check_num_Points()
        {
            if (this.Polyline)
            {
                if (this.Points.Length < 2)
                {
                    new System.ArgumentOutOfRangeException("Need at least 2 points for a polyline", nameof(this.Points));
                }
            }
            else if (this.Bezier)
            {
                if (this.Points.Length < 4)
                {
                    // two points
                    // two control points
                    new System.ArgumentOutOfRangeException("Need at least 4 points for a bezier", nameof(this.Points));
                }
            }
        }

        private void _drop_shape()
        {

            var targetpage = VisioScripting.TargetPage.Auto.ResolveToPage(this.Client);

            var shapeids = this.Client.Master.DropMasters(targetpage, this.Master, this.DropPosition);
            var shape_objects = VisioAutomation.Shapes.ShapeHelper.GetShapesFromIDs(targetpage.Page.Shapes, shapeids);

            // If there are cells to set, then use them
            if (this.Cells != null)
            {
                var writer = new VisioAutomation.ShapeSheet.Writers.SidSrcWriter();
                writer.BlastGuards = true;
                writer.TestCircular = true;

                for (int i = 0; i < shapeids.Count(); i++)
                {
                    var shapeid = shapeids[i];
                    var shape_cells = this.Cells[i % this.Cells.Length];

                    shape_cells.Apply(writer, (short)shapeid);
                }

                var surface = new VisioAutomation.SurfaceTarget(targetpage.Page);

                using (var undoscope = this.Client.Undo.NewUndoScope(nameof(NewVisioShape) +":CommitCells"))
                {
                    writer.CommitFormulas(surface);
                }

            }


            // Visio does not select dropped masters by default - unlike shapes that are directly drawn
            // so force visio to select the dropped shapes

            ((SMA.Cmdlet)this).WriteVerbose("Clearing the selection");
            this.Client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);
            ((SMA.Cmdlet)this).WriteVerbose("Selecting the shapes that were dropped");
            this.Client.Selection.SelectShapes(VisioScripting.TargetWindow.Auto, shape_objects);

            this.WriteObject(shape_objects, true);
        }
    }
}