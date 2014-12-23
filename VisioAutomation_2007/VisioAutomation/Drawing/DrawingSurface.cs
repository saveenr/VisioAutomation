using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Drawing
{
    public struct DrawingSurface
    {
        public IVisio.Page Page;
        public IVisio.Master Master;
        public IVisio.Shape Shape;

        public DrawingSurface(IVisio.Page page)
        {
            this.Page = page;
            this.Master = null;
            this.Shape = null;
        }

        public DrawingSurface(IVisio.Master master)
        {
            this.Page = null;
            this.Master = master;
            this.Shape = null;
        }


        public DrawingSurface(IVisio.Shape shape)
        {
            this.Page = null;
            this.Master = null;
            this.Shape = shape;
        }

        public IVisio.Shape DrawLine(VA.Drawing.Point p1, VA.Drawing.Point p2)
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
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }

        }

        public IVisio.Shape DrawPolyLine(IList<VA.Drawing.Point> points)
        {
            var doubles_array = VA.Drawing.Point.ToDoubles(points).ToArray();

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
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }
        }

        public IVisio.Shape DrawBezier(IList<VA.Drawing.Point> points, short degree, short flags)
        {
            var doubles_array = VA.Drawing.Point.ToDoubles(points).ToArray();
 
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
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }
        }

        public IVisio.Shape DrawBezier(IList<VA.Drawing.Point> points)
        {
            short degree = 3;
            short flags = 0;
            var shape = this.DrawBezier(points, degree, flags);
            return shape;
        }

        public IVisio.Shape DrawOval(VA.Drawing.Rectangle rect)
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
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }
        }

        public IVisio.Shape DrawOval(VA.Drawing.Point center, double radius)
        {
            var A = center.Add(-radius, -radius);
            var B = center.Add(radius, radius);
            var rect = new VA.Drawing.Rectangle(A, B);

            return this.DrawOval(rect);
        }

        public IVisio.Shape DrawRectangle(VA.Drawing.Rectangle rect)
        {
            var shape = this.DrawRectangle(rect.Left, rect.Bottom, rect.Right, rect.Top);
            return shape;
        }

        public IVisio.Shape DrawRectangle(double x0, double y0, double x1, double y1)
        {
            if (this.Master != null)
            {
                var shape = this.Master.DrawRectangle(x0,y0,x1,y1);
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
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }
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
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }
        }

        public IVisio.Shape DrawNURBS(IList<VA.Drawing.Point> controlpoints,
                                     IList<double> knots,
                                     IList<double> weights, int degree)
        {
            // flags:
            // None = 0,
            // IVisio.VisDrawSplineFlags.visSpline1D

            var flags = 0;
            double[] pts_dbl_a = VA.Drawing.Point.ToDoubles(controlpoints).ToArray();
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
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }

        }

        public short[] DropManyU(
            IList<IVisio.Master> masters,
            IEnumerable<VA.Drawing.Point> points)
        {
            if (masters == null)
            {
                throw new System.ArgumentNullException("masters");
            }

            if (masters.Count < 1)
            {
                return new short[0];
            }

            if (points == null)
            {
                throw new System.ArgumentNullException("points");
            }

            // NOTE: DropMany will fail if you pass in zero items to drop
            var masters_obj_array = masters.Cast<object>().ToArray();
            var xy_array = VA.Drawing.Point.ToDoubles(points).ToArray();

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

            short[] outids = (short[])outids_sa;
            return outids;
        }

        public IVisio.Shape Drop(
            IVisio.Master master,
            VA.Drawing.Point point)
        {
            if (master == null)
            {
                throw new System.ArgumentNullException("master");
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
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }
        }

        public IVisio.Shape DrawQuarterArc(VA.Drawing.Point p0, VA.Drawing.Point p1, IVisio.VisArcSweepFlags flags)
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
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }
        }



        public VA.Drawing.Rectangle GetBoundingBox(IVisio.VisBoundingBoxArgs args)
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

            var r = new VA.Drawing.Rectangle(bbx0, bby0, bbx1, bby1);
            return r;
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

        public List<IVisio.Shape> GetAllShapes()
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
            
            var list = new List<IVisio.Shape>();
            list.AddRange( shapes.AsEnumerable() );

            return list;
        }

        private static int check_stream_size(short[] stream, int chunksize)
        {
            if ((chunksize != 3) && (chunksize != 4))
            {
                throw new VA.AutomationException("Chunksize must be 3 or 4");
            }

            int remainder = stream.Length % chunksize;

            if (remainder != 0)
            {
                string msg = string.Format("stream must have a multiple of {0} elements", chunksize);
                throw new VA.AutomationException(msg);
            }

            return stream.Length / chunksize;
        }

        public string[] GetFormulasU_4(short[] stream)
        {
            int numitems = check_stream_size(stream, 4);
            if (numitems < 1)
            {
                return new string[0];
            }

            System.Array formulas_sa = null;

            if (this.Master != null)
            {
                this.Master.GetFormulasU(stream, out formulas_sa);
            }
            else if (this.Page != null)
            {
                this.Page.GetFormulasU(stream, out formulas_sa);
            }
            else if (this.Shape != null)
            {
                this.Shape.GetFormulasU(stream, out formulas_sa);
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }

            var formulas = get_formulas_array(formulas_sa, numitems);
            return formulas;
        }

        public string[] GetFormulasU_3(short[] stream)
        {
            int numitems = check_stream_size(stream, 3);
            if (numitems < 1)
            {
                return new string[0];
            }

            System.Array formulas_sa = null;

            if (this.Master != null)
            {
                this.Master.GetFormulasU(stream, out formulas_sa);
            }
            else if (this.Page != null)
            {
                this.Page.GetFormulasU(stream, out formulas_sa);
            }
            else if (this.Shape != null)
            {
                this.Shape.GetFormulasU(stream, out formulas_sa);
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }

            var formulas = get_formulas_array(formulas_sa, numitems);
            return formulas;
        }

        private static string[] get_formulas_array(System.Array formulas_sa, int numitems)
        {
            object[] formulas_obj_array = (object[])formulas_sa;

            if (formulas_obj_array.Length != numitems)
            {
                string msg = string.Format(
                    "Expected {0} items from GetFormulas but only received {1}",
                    numitems,
                    formulas_obj_array.Length);
                throw new AutomationException(msg);
            }

            string[] formulas = new string[formulas_obj_array.Length];
            formulas_obj_array.CopyTo(formulas, 0);
            return formulas;
        }

        public TResult[] GetResults_4<TResult>(short[] stream, IList<IVisio.VisUnitCodes> unitcodes)
        {
            EnforceValidResultType(typeof(TResult));

            int numitems = check_stream_size(stream, 4);
            if (numitems < 1)
            {
                return new TResult[0];
            }

            var result_type = typeof(TResult);
            var unitcodes_obj_array = get_unit_code_obj_array(unitcodes);
            var flags = get_VisGetSetArgs(result_type);

            System.Array results_sa = null;

            if (this.Master != null)
            {
                this.Master.GetResults(stream, (short)flags, unitcodes_obj_array, out results_sa);
            }
            else if (this.Page != null)
            {
                this.Page.GetResults(stream, (short)flags, unitcodes_obj_array, out results_sa);
            }
            else if (this.Shape != null)
            {
                this.Shape.GetResults(stream, (short)flags, unitcodes_obj_array, out results_sa);
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }
            
            var results = get_results_array<TResult>(results_sa, numitems);
            return results;
        }

        public TResult[] GetResults_3<TResult>(short[] stream, IList<IVisio.VisUnitCodes> unitcodes)
        {
            EnforceValidResultType(typeof(TResult));

            int numitems = check_stream_size(stream, 3);
            if (numitems < 1)
            {
                return new TResult[0];
            }

            var result_type = typeof(TResult);
            var unitcodes_obj_array = get_unit_code_obj_array(unitcodes);
            var flags = get_VisGetSetArgs(result_type);

            System.Array results_sa = null;

            if (this.Master != null)
            {
                this.Master.GetResults(stream, (short)flags, unitcodes_obj_array, out results_sa);
            }
            else if (this.Page != null)
            {
                this.Page.GetResults(stream, (short)flags, unitcodes_obj_array, out results_sa);
            }
            else if (this.Shape != null)
            {
                this.Shape.GetResults(stream, (short)flags, unitcodes_obj_array, out results_sa);
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }

            var results = get_results_array<TResult>(results_sa, numitems);
            return results;
        }



        private static TResult[] get_results_array<TResult>(System.Array results_sa, int numitems)
        {
            if (results_sa.Length != numitems)
            {
                string msg = string.Format(
                    "Expected {0} items from GetResults but only received {1}",
                    numitems,
                    results_sa.Length);
                throw new AutomationException(msg);
            }

            TResult[] results = new TResult[results_sa.Length];
            results_sa.CopyTo(results, 0);
            return results;
        }

        private static IVisio.VisGetSetArgs get_VisGetSetArgs(System.Type type)
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
                string msg = string.Format("Internal error: Unsupported Result Type: {0}", type.Name);
                throw new VA.AutomationException(msg);
            }
            return flags;
        }

        private static object[] get_unit_code_obj_array(IList<IVisio.VisUnitCodes> unitcodes)
        {
            // Create the unit codes array
            object[] unitcodes_obj_array = null;
            if (unitcodes != null)
            {
                unitcodes_obj_array = new object[unitcodes.Count];
                for (int i = 0; i < unitcodes.Count; i++)
                {
                    unitcodes_obj_array[i] = unitcodes[i];
                }
            }
            return unitcodes_obj_array;
        }

        internal static void EnforceValidResultType(System.Type result_type)
        {
            if (!IsValidResultType(result_type))
            {
                string msg = string.Format("Unsupported Result Type: {0}", result_type.Name);
                throw new VA.AutomationException(msg);
            }
        }

        internal static bool IsValidResultType(System.Type result_type)
        {
            return (result_type == typeof(int)
                    || result_type == typeof(double)
                    || result_type == typeof(string));
        }

        public int SetFormulas(short [] stream, object[] formulas, short flags)
        {
            int c;
            if (this.Shape != null)
            {
                c = this.Shape.SetFormulas(stream, formulas, flags);
            }
            else if (this.Master != null)
            {
                c = this.Master.SetFormulas(stream, formulas, flags);
            }
            else if (this.Page != null)
            {
                c = this.Page.SetFormulas(stream, formulas, flags);
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");                
            }

            return c;
        }

        public int SetResults(short[] stream, object[] unitcodes, object[] results, short flags)
        {
            int c;
            if (this.Shape != null)
            {
                c = this.Shape.SetResults(stream, unitcodes,results, flags);
            }
            else if (this.Master != null)
            {
                c = this.Master.SetResults(stream, unitcodes, results, flags);
            }
            else if (this.Page != null)
            {
                c = this.Page.SetResults(stream, unitcodes, results, flags);
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }

            return c;
        }

    }
}