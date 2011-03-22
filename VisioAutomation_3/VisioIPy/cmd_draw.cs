using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;
using System.Linq;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        public IVisio.Shape DrawLine(double x0, double y0, double x1, double y1)
        {
            return this.ScriptingSession.Draw.DrawLine(x0, y0, x1, y1);
        }

        public IVisio.Shape DrawLine(VA.Drawing.Point p0, VA.Drawing.Point p1)
        {
            return this.ScriptingSession.Draw.DrawLine(p0.X, p0.Y, p1.X, p1.Y);
        }

        public IVisio.Shape DrawPolyLine(IVisio.Application app, IList<VA.Drawing.Point> points)
        {
            return this.ScriptingSession.Draw.DrawPolyLine(points);
        }

        public IVisio.Shape DrawPolyLine(params VA.Drawing.Point[] points)
        {
            return this.ScriptingSession.Draw.DrawPolyLine(points);
        }

        public IVisio.Shape DrawPolyLine(IList<VA.Drawing.Point> points)
        {
            return this.ScriptingSession.Draw.DrawPolyLine(points);
        }

        public IVisio.Shape DrawPolyLine(params double[] doubles)
        {
            return this.ScriptingSession.Draw.DrawPolyLine(VA.Drawing.DrawingUtil.DoublesToPoints(doubles).ToList());
        }

        public IVisio.Shape DrawPolyLine(IList<double> doubles)
        {
            return this.ScriptingSession.Draw.DrawPolyLine(VA.Drawing.DrawingUtil.DoublesToPoints(doubles).ToList());
        }

        public IVisio.Shape DrawNURBS(IList<VA.Drawing.Point> controlpoints, IList<double> knots,
                                      IList<double> weights,
                                      int degree)
        {
            return this.ScriptingSession.Draw.DrawNURBSCurve(controlpoints, knots, weights, degree);
        }

        public IVisio.Shape DrawOval(double x0, double y0, double x1, double y1)
        {
            return this.ScriptingSession.Draw.DrawOval(x0, y0, x1, y1);
        }

        public IVisio.Shape DrawOval(VA.Drawing.Point origin, double radius)
        {
            return this.ScriptingSession.Draw.DrawOval(origin, radius);
        }

        public IVisio.Shape DrawPieSlice(double cx, double cy,
                                         double radius,
                                         double start_angle,
                                         double end_angle)
        {
            return this.ScriptingSession.Draw.DrawPieSlice(new VA.Drawing.Point(cx, cy), radius,
                                                 start_angle, end_angle);
        }

        public IList<IVisio.Shape> DrawPieSlices(double cx, double cy, double radius, IList<double> values)
        {
            return this.ScriptingSession.Draw.DrawPieSlices(new VA.Drawing.Point(cx, cy), radius,
                                                  values);
        }

        public IVisio.Shape DrawRectangle(double x0, double y0, double x1, double y1)
        {
            return this.ScriptingSession.Draw.DrawRectangle(x0, y0, x1, y1);
        }

        public IVisio.Shape DrawRectangle(VA.Drawing.Point p0, VA.Drawing.Point p1)
        {
            return this.ScriptingSession.Draw.DrawRectangle(p0.X, p0.Y, p1.X, p1.Y);
        }

        public IVisio.Shape DrawRectangle(VA.Drawing.Rectangle rect)
        {
            return DrawRectangle(rect.LowerLeft, rect.UpperRight);
        }

        public IVisio.Shape DrawBezier(IList<VA.Drawing.Point> points)
        {
            return this.ScriptingSession.Draw.DrawBezier(points);
        }

        public IVisio.Shape DrawBezier(IList<double> doubles)
        {
            return this.ScriptingSession.Draw.DrawBezier(VA.Drawing.DrawingUtil.DoublesToPoints(doubles));
        }

        public VA.Drawing.Size DefaultCellSize = new VA.Drawing.Size(1.0, 0.5);

        public IList<short> DrawGrid(IVisio.Master master, double cw, double ch, int cols, int rows)
        {
            var cellsize = new VA.Drawing.Size(cw, ch);
            return this.ScriptingSession.Draw.DrawGrid(master, cellsize, cols, rows);
        }
    }
}