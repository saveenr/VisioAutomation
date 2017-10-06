using System.Collections.Generic;
using System.Linq;
using SMA = System.Management.Automation;
using VisioAutomation.Geometry;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioShape)]
    public class NewVisioShape : VisioCmdlet
    {
        [SMA.Parameter(ParameterSetName = "masters", Position = 0, Mandatory = true)]
        public IVisio.Master[] Masters { get; set; }

        [SMA.Parameter(ParameterSetName = "shape", Position = 0, Mandatory = true)]
        public ShapeType Type { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public double [] Points { get; set; }

        [SMA.Parameter(ParameterSetName = "masters", Mandatory = false)]
        public string[] Names { get; set; }

        [SMA.Parameter(ParameterSetName = "masters", Mandatory = false)]
        public SMA.SwitchParameter NoSelect=false;

        protected override void ProcessRecord()
        {
            if (this.Masters != null)
            {
                drop_shape();
            }
            else
            {
                draw_shape();
            }
        }

        private void draw_shape()
        {
            var points = VisioAutomation.Geometry.Point.FromDoubles(this.Points).ToList();

            check_points_for_shape_type(points);

            if (this.Type == ShapeType.Rectangle)
            {
                var r = new VisioAutomation.Geometry.Rectangle(points[0], points[1]);
                var shape = this.Client.Draw.Rectangle(r);
                this.WriteObject(shape);
            }
            else if (this.Type == ShapeType.Line)
            {
                var lineseg = new VisioAutomation.Geometry.LineSegment(points[0], points[1]);
                var shape = this.Client.Draw.Line(lineseg.Start, lineseg.End);
                this.WriteObject(shape);
            }
            else if (this.Type == ShapeType.Oval)
            {
                var r = new VisioAutomation.Geometry.Rectangle(points[0], points[1]);
                var shape = this.Client.Draw.Oval(r);
                this.WriteObject(shape);
            }
            else if (this.Type == ShapeType.Polyline)
            {
                var shape = this.Client.Draw.PolyLine(points);
                this.WriteObject(shape);
            }
            else if (this.Type == ShapeType.Bezier)
            {
                var shape = this.Client.Draw.Bezier(points);
                this.WriteObject(shape);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }
        }

        private void check_points_for_shape_type(List<Point> points)
        {
            if (this.Type == ShapeType.Rectangle || this.Type == ShapeType.Line || this.Type == ShapeType.Oval)
            {
                if (points.Count != 2)
                {
                    string msg = string.Format("Need 2 points for a {0}", this.Type);
                    new System.ArgumentOutOfRangeException(msg);
                }
            }
            else if(this.Type == ShapeType.Polyline)
            {
                if (points.Count < 2)
                {
                    new System.ArgumentOutOfRangeException("Need at leat 2 points for a polyline");
                }
            }
            else if (this.Type == ShapeType.Bezier)
            {
                if (points.Count < 2)
                {
                    new System.ArgumentOutOfRangeException("Need at leat 2 points for a bezier");
                }
            }
        }

        private void drop_shape()
        {
            this.WriteVerbose("NoSelect: {0}", this.NoSelect);

            var points = VisioAutomation.Geometry.Point.FromDoubles(this.Points).ToList();
            var shape_ids = this.Client.Master.Drop(this.Masters, points);

            var page = this.Client.Page.Get();
            var shape_objects = VisioAutomation.Shapes.ShapeHelper.GetShapesFromIDs(page.Shapes, shape_ids);

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

            this.Client.Selection.SelectNone();

            if (!this.NoSelect)
            {
                // Select the Shapes
                ((SMA.Cmdlet) this).WriteVerbose("Selecting");
                this.Client.Selection.Select(shape_objects);
            }
            this.WriteObject(shape_objects, false);
        }
    }
}