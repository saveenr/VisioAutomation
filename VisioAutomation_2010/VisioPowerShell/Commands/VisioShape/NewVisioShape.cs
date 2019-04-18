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
        public IVisio.Master[] Masters { get; set; }

        [SMA.Parameter(ParameterSetName = "shape", Position = 0, Mandatory = true)]
        public Models.ShapeType Type { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public double [] Points { get; set; }

        [SMA.Parameter(ParameterSetName = "masters", Mandatory = false)]
        public string[] Names { get; set; }

        [SMA.Parameter(ParameterSetName = "masters", Mandatory = false)]
        public VisioPowerShell.Models.ShapeCells[] Cells { get; set; }

        [SMA.Parameter(ParameterSetName = "masters", Mandatory = false)]
        public SMA.SwitchParameter NoSelect=false;

        protected override void ProcessRecord()
        {
            if (this.Masters != null)
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
            var points = VisioAutomation.Geometry.Point.FromDoubles(this.Points).ToList();

            _check_points_for_shape_type(points);

            if (this.Type == Models.ShapeType.Rectangle)
            {
                var r = new VisioAutomation.Geometry.Rectangle(points[0], points[1]);
                var shape = this.Client.Draw.DrawRectangle(r);
                this.WriteObject(shape);
            }
            else if (this.Type == Models.ShapeType.Line)
            {
                var lineseg = new VisioAutomation.Models.Geometry.LineSegment(points[0], points[1]);
                var shape = this.Client.Draw.DrawLine(lineseg.Start, lineseg.End);
                this.WriteObject(shape);
            }
            else if (this.Type == Models.ShapeType.Oval)
            {
                var r = new VisioAutomation.Geometry.Rectangle(points[0], points[1]);
                var shape = this.Client.Draw.DrawOval(r);
                this.WriteObject(shape);
            }
            else if (this.Type == Models.ShapeType.Polyline)
            {
                var shape = this.Client.Draw.DrawPolyLine(points);
                this.WriteObject(shape);
            }
            else if (this.Type == Models.ShapeType.Bezier)
            {
                var shape = this.Client.Draw.DrawBezier(points);
                this.WriteObject(shape);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }
        }

        private void _check_points_for_shape_type(List<VisioAutomation.Geometry.Point> points)
        {
            if (this.Type == Models.ShapeType.Rectangle || this.Type == Models.ShapeType.Line || this.Type == Models.ShapeType.Oval)
            {
                if (points.Count != 2)
                {
                    string msg = string.Format("Need 2 points for a {0}", this.Type);
                    new System.ArgumentOutOfRangeException(msg);
                }
            }
            else if(this.Type == Models.ShapeType.Polyline)
            {
                if (points.Count < 2)
                {
                    new System.ArgumentOutOfRangeException("Need at leat 2 points for a polyline", nameof(points));
                }
            }
            else if (this.Type == Models.ShapeType.Bezier)
            {
                if (points.Count < 2)
                {
                    new System.ArgumentOutOfRangeException("Need at leat 2 points for a bezier", nameof(points));
                }
            }
        }

        private void _drop_shape()
        {
            this.WriteVerbose("NoSelect: {0}", this.NoSelect);

            var points = VisioAutomation.Geometry.Point.FromDoubles(this.Points).ToList();

            var shapeids = this.Client.Master.DropMasters(VisioScripting.TargetPage.Auto, this.Masters, points);

            var page = this.Client.Page.GetActivePage();
            var shape_objects = VisioAutomation.Shapes.ShapeHelper.GetShapesFromIDs(page.Shapes, shapeids);

            // If Names is not empty... assign it to the shape
            if (this.Names != null)
            {
                int up_to = System.Math.Min(shape_objects.Count, this.Names.Length);
                for (int i = 0; i < up_to; i++)
                {
                    string cur_name = this.Names[i];
                    if (cur_name != null)
                    {
                        var cur_shape = shape_objects[i];
                        cur_shape.NameU = cur_name;
                    }
                }
            }

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

                var surface = this.Client.ShapeSheet.GetShapeSheetSurface();

                using (var undoscope = this.Client.Undo.NewUndoScope(nameof(NewVisioShape) +":CommitCells"))
                {
                    writer.CommitFormulas(surface);
                }

            }

            this.Client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);

            if (!this.NoSelect)
            {
                // Select the Shapes
                ((SMA.Cmdlet)this).WriteVerbose("Selecting");
                this.Client.Selection.SelectShapes(VisioScripting.TargetWindow.Auto, shape_objects);
            }

            this.WriteObject(shape_objects, true);
        }
    }
}