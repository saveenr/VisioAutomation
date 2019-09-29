using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation
{
    public enum SurfaceTargetType
    {
        Page,
        Master,
        Shape
    }
    public struct SurfaceTarget
    {
        public readonly IVisio.Page Page;
        public readonly IVisio.Master Master;
        public readonly IVisio.Shape Shape;
        public readonly SurfaceTargetType TargetType;

        public SurfaceTarget(IVisio.Page page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            this.Page = page;
            this.Master = null;
            this.Shape = null;

            this.TargetType = SurfaceTargetType.Page;
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

            this.TargetType = SurfaceTargetType.Master;

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

            this.TargetType = SurfaceTargetType.Shape;
        }

        public IVisio.Shapes Shapes
        {
            get
            {
                var shapes = this.TargetType switch
                {
                    SurfaceTargetType.Master => this.Master.Shapes,
                    SurfaceTargetType.Page => this.Page.Shapes,
                    SurfaceTargetType.Shape => this.Shape.Shapes,
                    _ => throw new System.ArgumentException("Unhandled Drawing Surface")
                };

                return shapes;
            }
        }

        public short ID16
        {
            get
            {
                short id16 = this.TargetType switch
                {
                    SurfaceTargetType.Master => this.Master.ID16,
                    SurfaceTargetType.Page => this.Page.ID16,
                    SurfaceTargetType.Shape => this.Shape.ID16,
                    _ => throw new System.ArgumentException("Unhandled Drawing Surface")
                };

                return id16;
            }
        }

        public IVisio.Shape DrawLine(VisioAutomation.Geometry.Point p1, VisioAutomation.Geometry.Point p2)
        {
            var shape= this.TargetType switch
            {
                SurfaceTargetType.Master => this.Master.DrawLine(p1.X, p1.Y, p2.X, p2.Y),
                SurfaceTargetType.Page => this.Page.DrawLine(p1.X, p1.Y, p2.X, p2.Y),
                SurfaceTargetType.Shape => this.Shape.DrawLine(p1.X, p1.Y, p2.X, p2.Y),
                _ => throw new System.ArgumentException("Unhandled Drawing Surface")
            };

            return shape;
        }

        public IVisio.Shape DrawPolyLine(IList<VisioAutomation.Geometry.Point> points)
        {
            var doubles_array = VisioAutomation.Geometry.Point.ToDoubles(points).ToArray();

            var shape = this.TargetType switch
            {
                SurfaceTargetType.Master => this.Master.DrawPolyline(doubles_array, 0),
                SurfaceTargetType.Page => this.Page.DrawPolyline(doubles_array, 0),
                SurfaceTargetType.Shape => this.Shape.DrawPolyline(doubles_array, 0),
                _ => throw new System.ArgumentException("Unhandled Drawing Surface")
            };

            return shape;

        }

        public IVisio.Shape DrawBezier(IList<VisioAutomation.Geometry.Point> points, short degree, short flags)
        {
            var doubles_array = VisioAutomation.Geometry.Point.ToDoubles(points).ToArray();

            var shape = this.TargetType switch
            {
                SurfaceTargetType.Master => this.Master.DrawBezier(doubles_array, degree, flags),
                SurfaceTargetType.Page => this.Page.DrawBezier(doubles_array, degree, flags),
                SurfaceTargetType.Shape => this.Shape.DrawBezier(doubles_array, degree, flags),
                _ => throw new System.ArgumentException("Unhandled Drawing Surface")
            };

            return shape;


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
            var shape = this.TargetType switch
            {
                SurfaceTargetType.Master => this.Master.DrawOval(rect.Left, rect.Bottom, rect.Right, rect.Top),
                SurfaceTargetType.Page => this.Page.DrawOval(rect.Left, rect.Bottom, rect.Right, rect.Top),
                SurfaceTargetType.Shape => this.Shape.DrawOval(rect.Left, rect.Bottom, rect.Right, rect.Top),
                _ => throw new System.ArgumentException("Unhandled Drawing Surface")
            };

            return shape;
        }

        public IVisio.Shape DrawRectangle(VisioAutomation.Geometry.Rectangle rect)
        {
            var shape = this.DrawRectangle(rect.Left, rect.Bottom, rect.Right, rect.Top);
            return shape;
        }

        public IVisio.Shape DrawRectangle(double x0, double y0, double x1, double y1)
        {

            var shape = this.TargetType switch
            {
                SurfaceTargetType.Master => this.Master.DrawRectangle(x0, y0, x1, y1),
                SurfaceTargetType.Page => this.Page.DrawRectangle(x0, y0, x1, y1),
                SurfaceTargetType.Shape => this.Shape.DrawRectangle(x0, y0, x1, y1),
                _ => throw new System.ArgumentException("Unhandled Drawing Surface")
            };

            return shape;
        }

        public IVisio.Shape DrawLine(double x0, double y0, double x1, double y1)
        {
            var shape = this.TargetType switch
            {
                SurfaceTargetType.Master => this.Master.DrawLine(x0, y0, x1, y1),
                SurfaceTargetType.Page => this.Page.DrawLine(x0, y0, x1, y1),
                SurfaceTargetType.Shape => this.Shape.DrawLine(x0, y0, x1, y1),
                _ => throw new System.ArgumentException("Unhandled Drawing Surface")
            };

            return shape;
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


            var shape = this.TargetType switch
            {
                SurfaceTargetType.Master => this.Master.DrawNURBS((short)degree, (short)flags, pts_dbl_a, kts_dbl_a, weights_dbl_a),
                SurfaceTargetType.Page => this.Page.DrawNURBS((short)degree, (short)flags, pts_dbl_a, kts_dbl_a, weights_dbl_a),
                SurfaceTargetType.Shape => this.Shape.DrawNURBS((short)degree, (short)flags, pts_dbl_a, kts_dbl_a, weights_dbl_a),
                _ => throw new System.ArgumentException("Unhandled Drawing Surface")
            };

            return shape;
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

            var val = this.TargetType switch
            {
                SurfaceTargetType.Master => this.Master.DropManyU(masters_obj_array, xy_array, out outids_sa),
                SurfaceTargetType.Page => this.Page.DropManyU(masters_obj_array, xy_array, out outids_sa),
                SurfaceTargetType.Shape => this.Shape.DropManyU(masters_obj_array, xy_array, out outids_sa),
                _ => throw new System.ArgumentException("Unhandled Drawing Surface")
            };

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

            var shape = this.TargetType switch
            {
                SurfaceTargetType.Master => this.Master.Drop(master, point.X, point.Y),
                SurfaceTargetType.Page => this.Page.Drop(master, point.X, point.Y),
                SurfaceTargetType.Shape => this.Shape.Drop(master, point.X, point.Y),
                _ => throw new System.ArgumentException("Unhandled Drawing Surface")
            };

            return shape;

        }

        public IVisio.Shape DrawQuarterArc(VisioAutomation.Geometry.Point p0, VisioAutomation.Geometry.Point p1, IVisio.VisArcSweepFlags flags)
        {
            var shape = this.TargetType switch
            {
                SurfaceTargetType.Master => this.Master.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags),
                SurfaceTargetType.Page => this.Page.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags),
                SurfaceTargetType.Shape => this.Shape.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags),
                _ => throw new System.ArgumentException("Unhandled Drawing Surface")
            };

            return shape;
        }

        public VisioAutomation.Geometry.Rectangle GetBoundingBox(IVisio.VisBoundingBoxArgs args)
        {
            double bbx0, bby0, bbx1, bby1;

            switch (this.TargetType)
            {
                case SurfaceTargetType.Master:
                    this.Master.BoundingBox((short)args, out bbx0, out bby0, out bbx1, out bby1);
                    break;
                case SurfaceTargetType.Page:
                    this.Page.BoundingBox((short)args, out bbx0, out bby0, out bbx1, out bby1);
                    break;
                case SurfaceTargetType.Shape:
                    this.Shape.BoundingBox((short)args, out bbx0, out bby0, out bbx1, out bby1);
                    break;
                default:
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

            var val = this.TargetType switch
            {
                SurfaceTargetType.Master => this.Master.SetFormulas(stream.Array, formulas, flags),
                SurfaceTargetType.Page => this.Page.SetFormulas(stream.Array, formulas, flags),
                SurfaceTargetType.Shape => this.Shape.SetFormulas(stream.Array, formulas, flags),
                _ => throw new System.ArgumentException("Unhandled Drawing Surface")
            };

            return val;
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

            var val = this.TargetType switch
            {
                SurfaceTargetType.Master => this.Master.SetResults(stream.Array, unitcodes, results, flags),
                SurfaceTargetType.Page => this.Page.SetResults(stream.Array, unitcodes, results, flags),
                SurfaceTargetType.Shape => this.Shape.SetResults(stream.Array, unitcodes, results, flags),
                _ => throw new System.ArgumentException("Unhandled Drawing Surface")
            };

            return val;
        }

        public TResult[] GetResults<TResult>(ShapeSheet.Streams.StreamArray stream, object[] unitcodes)
        {
            if (stream.Array.Length == 0)
            {
                return new TResult[0];
            }

            _enforce_valid_result_type(typeof(TResult));

            var val = this.TargetType switch
            {
                SurfaceTargetType.Master => this.Master.GetResults<TResult>(stream, unitcodes),
                SurfaceTargetType.Page => this.Page.GetResults<TResult>(stream, unitcodes),
                SurfaceTargetType.Shape => this.Shape.GetResults<TResult>(stream, unitcodes),
                _ => throw new System.ArgumentException("Unhandled Drawing Surface")
            };

            return val;
        }

        public string[] GetFormulasU(ShapeSheet.Streams.StreamArray stream)
        {
            if (stream.Array.Length == 0)
            {
                return new string[0];
            }

            var val = this.TargetType switch
            {
                SurfaceTargetType.Master => this.Master.GetFormulasU(stream),
                SurfaceTargetType.Page => this.Page.GetFormulasU(stream),
                SurfaceTargetType.Shape => this.Shape.GetFormulasU(stream),
                _ => throw new System.ArgumentException("Unhandled Drawing Surface")
            };

            return val;
        }

        internal static T[] system_array_to_typed_array<T>(System.Array results_sa)
        {
            var results = new T[results_sa.Length];
            results_sa.CopyTo(results, 0);
            return results;
        }

        private static void _enforce_valid_result_type(System.Type result_type)
        {
            if (!_is_valid_result_type(result_type))
            {
                string msg = string.Format("Unsupported Result Type: {0}", result_type.Name);
                throw new VisioAutomation.Exceptions.InternalAssertionException(msg);
            }
        }

        private static bool _is_valid_result_type(System.Type result_type)
        {
            return (result_type == typeof(int)
                    || result_type == typeof(double)
                    || result_type == typeof(string));
        }

        internal static IVisio.VisGetSetArgs _type_to_vis_get_set_args(System.Type type)
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
