using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation
{
    public struct SurfaceTarget
    {
        public readonly IVisio.Page Page;
        public readonly IVisio.Master Master;
        public readonly IVisio.Shape Shape;

        public SurfaceTarget(IVisio.Page page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            this.Page = page;
            this.Master = null;
            this.Shape = null;
        }

        public SurfaceTarget(IVisio.Master master)
        {
            if (master == null)
            {
                throw new System.ArgumentNullException(nameof(master));
            }

            this.Page = null;
            this.Master = master;
            this.Shape = null;
        }

        public SurfaceTarget(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            this.Page = null;
            this.Master = null;
            this.Shape = shape;
        }

        public IVisio.Shapes Shapes
        {
            get
            {
                IVisio.Shapes shapes;

                if (this.Master != null)
                {
                    shapes = this.Master.Shapes;
                }
                else if (this.Page != null)
                {
                    shapes = this.Page.Shapes;
                }
                else if (this.Shape != null)
                {
                    shapes = this.Shape.Shapes;
                }
                else
                {
                    throw new System.ArgumentException("Unhandled Drawing Surface");
                }
                return shapes;
            }
        }

        public short ID16
        {
            get
            {
                if (this.Shape != null)
                {
                    return this.Shape.ID16;
                }
                else if (this.Page != null)
                {
                    return this.Page.ID16;
                }
                else if (this.Master != null)
                {
                    return this.Master.ID16;
                }
                else
                {
                    throw new System.ArgumentException("Unhandled Drawing Surface");
                }
            }
        }

        public IVisio.Shape DrawLine(VisioAutomation.Geometry.Point p1, VisioAutomation.Geometry.Point p2)
        {

            if (this.Master != null)
            {
                var shape = this.Master.DrawLine(p1.X, p1.Y, p2.X, p2.Y);
                return shape;
            }
            else if (this.Page != null)
            {
                var shape = this.Page.DrawLine(p1.X, p1.Y, p2.X, p2.Y);
                return shape;
            }
            else if (this.Shape != null)
            {
                var shape = this.Shape.DrawLine(p1.X, p1.Y, p2.X, p2.Y);
                return shape;
            }

            throw new System.ArgumentException("Unhandled Drawing Surface");

        }

        public IVisio.Shape DrawPolyLine(IList<VisioAutomation.Geometry.Point> points)
        {
            var doubles_array = VisioAutomation.Geometry.Point.ToDoubles(points).ToArray();

            if (this.Master != null)
            {
                var shape = this.Master.DrawPolyline(doubles_array, 0);
                return shape;
            }
            else if (this.Page != null)
            {
                var shape = this.Page.DrawPolyline(doubles_array, 0);
                return shape;
            }
            else if (this.Shape != null)
            {
                var shape = this.Shape.DrawPolyline(doubles_array, 0);
                return shape;
            }

            throw new System.ArgumentException("Unhandled Drawing Surface");
        }

        public IVisio.Shape DrawBezier(IList<VisioAutomation.Geometry.Point> points, short degree, short flags)
        {
            var doubles_array = VisioAutomation.Geometry.Point.ToDoubles(points).ToArray();

            if (this.Master != null)
            {
                var shape = this.Master.DrawBezier(doubles_array, degree, flags);
                return shape;
            }
            else if (this.Page != null)
            {
                var shape = this.Page.DrawBezier(doubles_array, degree, flags);
                return shape;
            }
            else if (this.Shape != null)
            {
                var shape = this.Shape.DrawBezier(doubles_array, degree, flags);
                return shape;
            }

            throw new System.ArgumentException("Unhandled Drawing Surface");

        }

        public IVisio.Shape DrawBezier(IList<VisioAutomation.Geometry.Point> points)
        {
            short degree = 3;
            short flags = 0;
            var shape = this.DrawBezier(points, degree, flags);
            return shape;
        }

        public IVisio.Shape DrawOval(VisioAutomation.Geometry.Rectangle rect)
        {
            if (this.Master != null)
            {
                var shape = this.Master.DrawOval(rect.Left, rect.Bottom, rect.Right, rect.Top);
                return shape;
            }
            else if (this.Page != null)
            {
                var shape = this.Page.DrawOval(rect.Left, rect.Bottom, rect.Right, rect.Top);
                return shape;
            }
            else if (this.Shape != null)
            {
                var shape = this.Shape.DrawOval(rect.Left, rect.Bottom, rect.Right, rect.Top);
                return shape;
            }

            throw new System.ArgumentException("Unhandled Drawing Surface");
        }

        public IVisio.Shape DrawRectangle(VisioAutomation.Geometry.Rectangle rect)
        {
            var shape = this.DrawRectangle(rect.Left, rect.Bottom, rect.Right, rect.Top);
            return shape;
        }

        public IVisio.Shape DrawRectangle(double x0, double y0, double x1, double y1)
        {
            if (this.Master != null)
            {
                var shape = this.Master.DrawRectangle(x0, y0, x1, y1);
                return shape;
            }
            else if (this.Page != null)
            {
                var shape = this.Page.DrawRectangle(x0, y0, x1, y1);
                return shape;
            }
            else if (this.Shape != null)
            {
                var shape = this.Shape.DrawRectangle(x0, y0, x1, y1);
                return shape;
            }

            throw new System.ArgumentException("Unhandled Drawing Surface");
            
        }

        public IVisio.Shape DrawLine(double x0, double y0, double x1, double y1)
        {
            if (this.Master != null)
            {
                var shape = this.Master.DrawLine(x0, y0, x1, y1);
                return shape;
            }
            else if (this.Page != null)
            {
                var shape = this.Page.DrawLine(x0, y0, x1, y1);

                return shape;
            }
            else if (this.Shape != null)
            {
                var shape = this.Shape.DrawLine(x0, y0, x1, y1);

                return shape;
            }

            throw new System.ArgumentException("Unhandled Drawing Surface");
            
        }

        public IVisio.Shape DrawNURBS(IList<VisioAutomation.Geometry.Point> controlpoints,
            IList<double> knots,
            IList<double> weights, int degree)
        {
            // flags:
            // None = 0,
            // IVisio.VisDrawSplineFlags.visSpline1D

            var flags = 0;
            double[] pts_dbl_a = VisioAutomation.Geometry.Point.ToDoubles(controlpoints).ToArray();
            double[] kts_dbl_a = knots.ToArray();
            double[] weights_dbl_a = weights.ToArray();

            if (this.Master != null)
            {
                var shape = this.Master.DrawNURBS((short)degree, (short)flags, pts_dbl_a, kts_dbl_a, weights_dbl_a);
                return shape;
            }
            else if (this.Page != null)
            {
                var shape = this.Page.DrawNURBS((short)degree, (short)flags, pts_dbl_a, kts_dbl_a, weights_dbl_a);
                return shape;
            }
            else if (this.Shape != null)
            {
                var shape = this.Shape.DrawNURBS((short)degree, (short)flags, pts_dbl_a, kts_dbl_a, weights_dbl_a);
                return shape;
            }

            throw new System.ArgumentException("Unhandled Drawing Surface");

        }

        public short[] DropManyU(
            IList<IVisio.Master> masters,
            IEnumerable<VisioAutomation.Geometry.Point> points)
        {
            if (masters == null)
            {
                throw new System.ArgumentNullException(nameof(masters));
            }

            if (masters.Count < 1)
            {
                return new short[0];
            }

            if (points == null)
            {
                throw new System.ArgumentNullException(nameof(points));
            }

            // NOTE: DropMany will fail if you pass in zero items to drop
            var masters_obj_array = masters.Cast<object>().ToArray();
            var xy_array = VisioAutomation.Geometry.Point.ToDoubles(points).ToArray();

            System.Array outids_sa;

            if (this.Master != null)
            {
                this.Master.DropManyU(masters_obj_array, xy_array, out outids_sa);
            }
            else if (this.Page != null)
            {
                this.Page.DropManyU(masters_obj_array, xy_array, out outids_sa);
            }
            else if (this.Shape != null)
            {
                this.Shape.DropManyU(masters_obj_array, xy_array, out outids_sa);
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }

            short[] outids = (short[]) outids_sa;
            return outids;
        }

        public IVisio.Shape Drop(
            IVisio.Master master,
            VisioAutomation.Geometry.Point point)
        {
            if (master == null)
            {
                throw new System.ArgumentNullException(nameof(master));
            }

            if (this.Master != null)
            {
                return this.Master.Drop(master, point.X, point.Y);
            }
            else if (this.Page != null)
            {
                return this.Page.Drop(master, point.X, point.Y);
            }
            else if (this.Shape != null)
            {
                return this.Shape.Drop(master, point.X, point.Y);
            }

            throw new System.ArgumentException("Unhandled Drawing Surface");
            
        }

        public IVisio.Shape DrawQuarterArc(VisioAutomation.Geometry.Point p0, VisioAutomation.Geometry.Point p1, IVisio.VisArcSweepFlags flags)
        {
            if (this.Master != null)
            {
                return this.Master.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags);
            }
            else if (this.Page != null)
            {
                return this.Page.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags);
            }
            else if (this.Shape != null)
            {
                return this.Shape.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags);
            }

            throw new System.ArgumentException("Unhandled Drawing Surface");
            
        }

        public VisioAutomation.Geometry.Rectangle GetBoundingBox(IVisio.VisBoundingBoxArgs args)
        {
            double bbx0, bby0, bbx1, bby1;
            if (this.Master != null)
            {
                this.Master.BoundingBox((short)args, out bbx0, out bby0, out bbx1, out bby1);
            }
            else if (this.Page != null)
            {
                this.Page.BoundingBox((short)args, out bbx0, out bby0, out bbx1, out bby1);
            }
            else if (this.Shape != null)
            {
                this.Shape.BoundingBox((short)args, out bbx0, out bby0, out bbx1, out bby1);
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }

            var r = new VisioAutomation.Geometry.Rectangle(bbx0, bby0, bbx1, bby1);
            return r;
        }

        public int SetFormulas(ShapeSheet.Streams.StreamArray stream, object[] formulas, short flags)
        {
            if (formulas.Length != stream.Count)
            {
                string msg =
                    string.Format("stream contains {0} items ({1} short values) and requires {2} formula values",
                        stream.Count, stream.Array.Length, stream.Count);
                throw new System.ArgumentException(msg);
            }

            if (this.Shape != null)
            {
                return this.Shape.SetFormulas(stream.Array, formulas, flags);
            }
            else if (this.Master != null)
            {
                return this.Master.SetFormulas(stream.Array, formulas, flags);
            }
            else if (this.Page != null)
            {
                return this.Page.SetFormulas(stream.Array, formulas, flags);
            }

            throw new System.ArgumentException("Unhandled Target");
        }

        public int SetResults(ShapeSheet.Streams.StreamArray stream, object[] unitcodes, object[] results, short flags)
        {
            if (results.Length != stream.Count)
            {
                string msg =
                    string.Format("stream contains {0} items ({1} short values) and requires {2} result values",
                        stream.Count, stream.Array.Length, stream.Count);
                throw new System.ArgumentException(msg);
            }

            if (this.Shape != null)
            {
                return this.Shape.SetResults(stream.Array, unitcodes, results, flags);
            }
            else if (this.Master != null)
            {
                return this.Master.SetResults(stream.Array, unitcodes, results, flags);
            }
            else if (this.Page != null)
            {
                return this.Page.SetResults(stream.Array, unitcodes, results, flags);
            }

            throw new System.ArgumentException("Unhandled Target");
        }

        public TResult[] GetResults<TResult>(ShapeSheet.Streams.StreamArray stream, object[] unitcodes)
        {
            if (stream.Array.Length == 0)
            {
                return new TResult[0];
            }

            EnforceValidResultType(typeof(TResult));

            var flags = TypeToVisGetSetArgs(typeof(TResult));

            System.Array results_sa = null;

            if (this.Master != null)
            {
                this.Master.GetResults(stream.Array, (short)flags, unitcodes, out results_sa);
            }
            else if (this.Page != null)
            {
                this.Page.GetResults(stream.Array, (short)flags, unitcodes, out results_sa);
            }
            else if (this.Shape != null)
            {
                this.Shape.GetResults(stream.Array, (short)flags, unitcodes, out results_sa);
            }
            else
            {
                throw new System.ArgumentException("Unhandled Target");
            }

            var results = system_array_to_typed_array<TResult>(results_sa);
            return results;
        }

        public string[] GetFormulasU(ShapeSheet.Streams.StreamArray stream)
        {
            if (stream.Array.Length == 0)
            {
                return new string[0];
            }

            System.Array formulas_sa = null;

            if (this.Master != null)
            {
                this.Master.GetFormulasU(stream.Array, out formulas_sa);
            }
            else if (this.Page != null)
            {
                this.Page.GetFormulasU(stream.Array, out formulas_sa);
            }
            else if (this.Shape != null)
            {
                this.Shape.GetFormulasU(stream.Array, out formulas_sa);
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }

            var formulas = system_array_to_typed_array<string>(formulas_sa);
            return formulas;
        }

        private static T[] system_array_to_typed_array<T>(System.Array results_sa)
        {
            var results = new T[results_sa.Length];
            results_sa.CopyTo(results, 0);
            return results;
        }

        private static void EnforceValidResultType(System.Type result_type)
        {
            if (!IsValidResultType(result_type))
            {
                string msg = string.Format("Unsupported Result Type: {0}", result_type.Name);
                throw new VisioAutomation.Exceptions.InternalAssertionException(msg);
            }
        }

        private static bool IsValidResultType(System.Type result_type)
        {
            return (result_type == typeof(int)
                    || result_type == typeof(double)
                    || result_type == typeof(string));
        }

        private static IVisio.VisGetSetArgs TypeToVisGetSetArgs(System.Type type)
        {
            IVisio.VisGetSetArgs flags;
            if (type == typeof(int))
            {
                flags = IVisio.VisGetSetArgs.visGetTruncatedInts;
            }
            else if (type == typeof(double))
            {
                flags = IVisio.VisGetSetArgs.visGetFloats;
            }
            else if (type == typeof(string))
            {
                flags = IVisio.VisGetSetArgs.visGetStrings;
            }
            else
            {
                string msg = string.Format("Unsupported Result Type: {0}", type.Name);
                throw new VisioAutomation.Exceptions.InternalAssertionException(msg);
            }
            return flags;
        }

    }
}
