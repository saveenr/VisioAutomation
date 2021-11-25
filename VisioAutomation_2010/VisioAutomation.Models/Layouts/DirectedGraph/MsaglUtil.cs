using System;
using System.Collections.Generic;
using MG = Microsoft.Msagl;

namespace VisioAutomation.Models.Layouts.DirectedGraph;

static class MsaglUtil
{
    public static VisioAutomation.Geometry.Rectangle ToVARectangle(MG.Core.Geometry.Rectangle n)
    {
        return new VisioAutomation.Geometry.Rectangle(n.Left, n.Bottom, n.Right, n.Top);
    }

    public static VisioAutomation.Geometry.Point ToVAPoint(MG.Core.Geometry.Point p)
    {
        return new VisioAutomation.Geometry.Point(p.X, p.Y);
    }

    public static IList<VisioAutomation.Geometry.Point> ToVAPoints(MG.Core.Layout.Edge edge)
    {

        if (edge.Curve is MG.Core.Geometry.Curves.Curve)
        {
            var curve = (MG.Core.Geometry.Curves.Curve)edge.Curve;

            var final_bez_points = new List<VisioAutomation.Geometry.Point> { MsaglUtil.ToVAPoint(edge.Curve.Start) };

            foreach (var cur_seg in curve.Segments)
            {
                if (cur_seg is MG.Core.Geometry.Curves.CubicBezierSegment)
                {
                    var bezier_seg = (MG.Core.Geometry.Curves.CubicBezierSegment)cur_seg;

                    // The first point at index 0 is deliberately not added
                    final_bez_points.Add(MsaglUtil.ToVAPoint(bezier_seg.B(1)));
                    final_bez_points.Add(MsaglUtil.ToVAPoint(bezier_seg.B(2)));
                    final_bez_points.Add(MsaglUtil.ToVAPoint(bezier_seg.B(3)));
                }
                else if (cur_seg is MG.Core.Geometry.Curves.LineSegment)
                {
                    var line_seg = (MG.Core.Geometry.Curves.LineSegment)cur_seg;
                    final_bez_points.Add(MsaglUtil.ToVAPoint(line_seg.Start));
                    final_bez_points.Add(MsaglUtil.ToVAPoint(line_seg.End));
                    final_bez_points.Add(MsaglUtil.ToVAPoint(line_seg.End));
                }
                else
                {
                    throw new InvalidOperationException("Unsupported Curve Segment type");
                }
            }

            return final_bez_points;
                
        }
        else if (edge.Curve is MG.Core.Geometry.Curves.LineSegment)
        {
            var final_bez_points = new List<VisioAutomation.Geometry.Point> { MsaglUtil.ToVAPoint(edge.Curve.Start) };
            var line_seg = (MG.Core.Geometry.Curves.LineSegment)edge.Curve;
            final_bez_points.Add(MsaglUtil.ToVAPoint(line_seg.Start));
            final_bez_points.Add(MsaglUtil.ToVAPoint(line_seg.End));
            final_bez_points.Add(MsaglUtil.ToVAPoint(line_seg.End));
            return final_bez_points;
                
        }

        throw new System.ArgumentException("Unhandled Curve Type");
    }
}