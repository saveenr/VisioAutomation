using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio= Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting
{
    static class DrawCommandsUtil
    {
        public static IList<IVisio.Shape> CreateDuplicates(IVisio.Page page,
                                                           IVisio.Shape shape,
                                                           int n)
        {
            // NOTE: n is the total number you want INCLUDING the original shape
            // example n=0 then result={s0}
            // example n=1, result={s0}
            // example n=2, result={s0,s1}
            // example n=3, result={s0,s1,s2}
            // where s0 is the original shape

            if (n < 2)
            {
                return new List<IVisio.Shape> {shape};
            }

            int num_doubles = (int) System.Math.Log(n, 2.0);
            int leftover = n - ((int)System.Math.Pow(2.0, num_doubles));
            if (leftover < 0)
            {
                throw new System.InvalidOperationException("internal error: leftover value must greater than or equal to zero");
            }

            var duplicated_shapes = new List<IVisio.Shape> {shape};

            var application = page.Application;
            var win = application.ActiveWindow;

            foreach (int i in Enumerable.Range(0, num_doubles))
            {
                win.DeselectAll();
                win.Select(duplicated_shapes, IVisio.VisSelectArgs.visSelect);
                var selection = win.Selection;
                selection.Duplicate();
                var selection1 = win.Selection;
                duplicated_shapes.AddRange(selection1.AsEnumerable());
            }

            if (leftover > 0)
            {
                var leftover_shapes = duplicated_shapes.Take(leftover);
                win.DeselectAll();
                win.Select(leftover_shapes, IVisio.VisSelectArgs.visSelect);
                var selection = win.Selection;
                selection.Duplicate();
                var selection1 = win.Selection;
                duplicated_shapes.AddRange(selection1.AsEnumerable());
            }

            win.DeselectAll();
            win.Select(duplicated_shapes, IVisio.VisSelectArgs.visSelect);

            if (duplicated_shapes.Count != n)
            {
                string msg = string.Format("internal error: failed to create {0} shapes, instead created {1}", n,
                                           duplicated_shapes.Count);
                throw new VA.Scripting.ScriptingException(msg);
            }

            var selection2 = win.Selection;
            if (selection2.Count != n)
            {
                throw new VA.Scripting.ScriptingException("internal error: failed to select the duplicated shapes");
            }

            return duplicated_shapes;
        }

        private static Drawing.Point GetPointAtRadius(Drawing.Point origin, double angle, double radius)
        {
            double x = radius*System.Math.Cos(angle);
            double y = radius*System.Math.Sin(angle);
            var new_point = new Drawing.Point(x,y);
            new_point = origin + new_point;
            return new_point;
        }

        public static IVisio.Shape DrawPieSlice(
            IVisio.Page page,
            VA.Drawing.Point center,
            double radius,
            double start_angle,
            double end_angle)
        {
            double total_angle = end_angle - start_angle;

            if (total_angle == 0.0)
            {
                // This devolves into a single line
                return page.DrawLine(center, GetPointAtRadius(center, start_angle, radius));
            }
            
            if (total_angle >= System.Math.PI*2.0)
            {
                // This devolves into a circle
                var A = center.Add(-radius, -radius);
                var B = center.Add(radius, radius);
                var rect = new VA.Drawing.Rectangle(A, B);
                var shape = page.DrawOval(rect);
                return shape;
            }

            // This is a true slice
            int degree;
            var sub_arcs = VA.Drawing.BezierSegment.FromArc(
                start_angle,
                end_angle);

            var arc_bez_points = (from p in VA.Drawing.BezierSegment.Merge(sub_arcs, out degree)
                                    select p.Multiply(radius) + center).ToList();

            var pie_points = new List<VA.Drawing.Point>();
            pie_points.Add(center);
            pie_points.Add(center);
            pie_points.Add(arc_bez_points[0]);
            pie_points.AddRange(arc_bez_points);
            pie_points.Add(arc_bez_points[arc_bez_points.Count - 1]);
            pie_points.Add(center);
            pie_points.Add(center);

            var doubles_array = VA.Drawing.Point.ToDoubles(pie_points).ToArray();
            var pie_slice = page.DrawBezier(doubles_array, (short)degree, 0);
            return pie_slice;
        }
    }
}