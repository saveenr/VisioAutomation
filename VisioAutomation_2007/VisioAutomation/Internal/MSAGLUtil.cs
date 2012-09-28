using System.Collections.Generic;
using System.Linq;
using MG = Microsoft.Msagl;
using VA=VisioAutomation;

namespace VisioAutomation.Internal
{
    static class MSAGLUtil
    {
        public static VA.Drawing.Rectangle ToVARectangle(MG.Splines.Rectangle n)
        {
            return new VA.Drawing.Rectangle(n.Left, n.Bottom, n.Right, n.Top);
        }

        public static VA.Drawing.Point ToVAPoint(MG.Point p)
        {
            return new VA.Drawing.Point(p.X, p.Y);
        }

        public static IList<VA.Drawing.Point> ToVAPoints(MG.Edge edge)
        {
            var final_bez_points = new List<VA.Drawing.Point> { ToVAPoint(edge.Curve.Start) };

            var curve = (MG.Splines.Curve) edge.Curve;

            foreach (var cur_seg in curve.Segments)
            {
                if (cur_seg is MG.Splines.CubicBezierSegment)
                {
                    var bezier_seg = (MG.Splines.CubicBezierSegment) cur_seg;

                    var bez_points =
                        new int[] { 0, 1, 2, 3 }
                            .Select(bezier_seg.B)
                            .Select(ToVAPoint)
                            .ToArray();

                    final_bez_points.AddRange(bez_points.Skip(1));
                }
                else if (cur_seg is MG.Splines.LineSegment)
                {
                    var line_seg = (MG.Splines.LineSegment) cur_seg;
                    final_bez_points.Add(ToVAPoint(line_seg.Start));
                    final_bez_points.Add(ToVAPoint(line_seg.End));
                    final_bez_points.Add(ToVAPoint(line_seg.End));
                }
                else
                {
                    throw new System.InvalidOperationException("Unsupported Curve Segment type");
                }
            }

            return final_bez_points;
        }
    }
}