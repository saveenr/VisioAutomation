using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio= Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting
{
    static class DrawCommandsUtil
    {
        private static Drawing.Rectangle[] GetInnerBorderRects(VA.Drawing.Rectangle R, double w)
        {
            VA.Drawing.Rectangle[] rects = {
                                    new VA.Drawing.Rectangle(R.LowerLeft.Add(0, 0), R.LowerLeft.Add(w, w)),
                                    new VA.Drawing.Rectangle(R.LowerRight.Add(-w, 0), R.LowerRight.Add(0, w)),
                                    new VA.Drawing.Rectangle(R.UpperLeft.Add(0, -w), R.UpperLeft.Add(w, 0)),
                                    new VA.Drawing.Rectangle(R.UpperRight.Add(-w, -w), R.UpperRight.Add(0, 0)),
                                    new VA.Drawing.Rectangle(R.LowerLeft.Add(0, w), R.UpperLeft.Add(w, -w)),
                                    new VA.Drawing.Rectangle(R.LowerRight.Add(-w, w), R.UpperRight.Add(0, -w)),
                                    new VA.Drawing.Rectangle(R.UpperLeft.Add(w, -w), R.UpperRight.Add(-w, 0)),
                                    new VA.Drawing.Rectangle(R.LowerLeft.Add(w, 0), R.LowerRight.Add(-w, w))
                                };
            return rects;
        }



        public static IVisio.Shape DrawInnerGlow(IVisio.Page page,
                                                 VA.Drawing.Rectangle rect,
                                                 double glow_width,
                                                 VA.Drawing.ColorRGB glow_color,
                                                 int glow_trans)
        {
            double bg_trans = 1.0;
            var bg_color = glow_color;

            var rects = GetInnerBorderRects(rect, glow_width);

            VA.Format.FillPattern[] grads = {
                                      VA.Format.FillPattern.RadialUpperRight,
                                      VA.Format.FillPattern.RadialUpperLeft,
                                      VA.Format.FillPattern.RadialLowerRight,
                                      VA.Format.FillPattern.RadialLowerLeft,
                                      VA.Format.FillPattern.LinearRightToLeft,
                                      VA.Format.FillPattern.LinearLeftToRight,
                                      VA.Format.FillPattern.LinearBottomToTop,
                                      VA.Format.FillPattern.LinearTopToBottom
                                  };

            var shapes = (from r in rects
                          select page.DrawRectangle(r))
                .ToArray();


            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

            for (int i=0;i<shapes.Length;i++)
            {
                short shapeid = (short) shapes[i].ID;
                update.SetFormula(shapeid, VA.ShapeSheet.SRCConstants.FillPattern, (int)grads[i]);
                update.SetFormula(shapeid, VA.ShapeSheet.SRCConstants.FillForegnd, VA.Convert.ColorToFormulaRGB(glow_color));
                update.SetFormula(shapeid, VA.ShapeSheet.SRCConstants.FillBkgnd, VA.Convert.ColorToFormulaRGB(bg_color));
                update.SetFormula(shapeid, VA.ShapeSheet.SRCConstants.FillForegndTrans, bg_trans);
                update.SetFormula(shapeid, VA.ShapeSheet.SRCConstants.FillBkgndTrans, glow_trans);
                update.SetFormula(shapeid, VA.ShapeSheet.SRCConstants.LinePattern, 0);
            }

            update.Execute(page);

            var application = page.Application;
            var active_window = application.ActiveWindow;
            active_window.DeselectAll();
            var group = VA.SelectionHelper.SelectAndGroup(active_window, shapes);
            VA.ShapeHelper.SetGroupSelectMode(group, IVisio.VisCellVals.visGrpSelModeGroupOnly);

            return group;
        }


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

            var duplicated_shapes = new List<IVisio.Shape>();
            duplicated_shapes.Add(shape);

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
                throw new AutomationException(msg);
            }

            var selection2 = win.Selection;
            if (selection2.Count != n)
            {
                throw new AutomationException("internal error: failed to select the duplicated shapes");
            }

            return duplicated_shapes;
        }

        public static IVisio.Shape DrawArcByThreePoints(
            IVisio.Page page,
            VA.Drawing.Point begin,
            VA.Drawing.Point end,
            VA.Drawing.Point control)
        {
            var shape = page.DrawArcByThreePoints(begin.X, begin.Y, end.X, end.Y, control.X, control.Y);
            return shape;
        }

        private static Drawing.Point GetPointAtRadius(Drawing.Point origin, double angle, double radius)
        {
            var new_point = new Drawing.Point(radius * System.Math.Cos(angle),
                                      radius * System.Math.Sin(angle));
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
                return page.DrawLine(center, GetPointAtRadius(center, start_angle, radius));
            }
            else if (total_angle >= 360)
            {
                var A = center.Add(-radius, -radius);
                var B = center.Add(radius, radius);
                var rect = new VA.Drawing.Rectangle(A, B);
                var shape = page.DrawOval(rect);
                return shape;
            }
            else
            {
                int degree;
                var sub_arcs = VA.Drawing.BezierSegment.FromArc(
                    Convert.DegreesToRadians(start_angle),
                    Convert.DegreesToRadians(end_angle));

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

                var doubles_array = VA.Drawing.DrawingUtil.PointsToDoubles(pie_points).ToArray();
                var pie_slice = page.DrawBezier(doubles_array, (short)degree, 0);
                return pie_slice;
            }
        }
    }
}