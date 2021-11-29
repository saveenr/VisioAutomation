using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet.Streams;

namespace VisioAutomation.Core
{
    public readonly struct VisioObjectTarget
    {
        public readonly IVisio.Page Page;
        public readonly IVisio.Master Master;
        public readonly IVisio.Shape Shape;
        public readonly VisioObjectCategory Category;
        private static readonly string _unhandled_category_exc_msg = "Unhandled Category";

        public VisioObjectTarget(IVisio.Page page)
        {
            this.Page = page ?? throw new System.ArgumentNullException(nameof(page));
            this.Master = null;
            this.Shape = null;
            this.Category = VisioObjectCategory.Page;
        }

        public VisioObjectTarget(IVisio.Master master)
        {
            this.Page = null;
            this.Master = master ?? throw new System.ArgumentNullException(nameof(master));
            this.Shape = null;
            this.Category = VisioObjectCategory.Master;
        }

        public VisioObjectTarget(IVisio.Shape shape)
        {
            this.Page = null;
            this.Master = null;
            this.Shape = shape ?? throw new System.ArgumentNullException(nameof(shape));
            this.Category = VisioObjectCategory.Shape;
        }

        public IVisio.Shapes Shapes
        {
            get
            {
                var shapes = this.Category switch
                {
                    VisioObjectCategory.Master => this.Master.Shapes,
                    VisioObjectCategory.Page => this.Page.Shapes,
                    VisioObjectCategory.Shape => this.Shape.Shapes,
                    _ => throw new System.ArgumentException(_unhandled_category_exc_msg)
                };

                return shapes;
            }
        }

        public short ID16
        {
            get
            {
                short id16 = this.Category switch
                {
                    VisioObjectCategory.Master => this.Master.ID16,
                    VisioObjectCategory.Page => this.Page.ID16,
                    VisioObjectCategory.Shape => this.Shape.ID16,
                    _ => throw new System.ArgumentException(_unhandled_category_exc_msg)
                };

                return id16;
            }
        }


        public short[] DropManyU(
            IList<IVisio.Master> masters,
            IEnumerable<Point> points)
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
            var xy_array = Point.ToDoubles(points).ToArray();

            System.Array outids_sa;

            var val = this.Category switch
            {
                VisioObjectCategory.Master => this.Master.DropManyU(masters_obj_array, xy_array, out outids_sa),
                VisioObjectCategory.Page => this.Page.DropManyU(masters_obj_array, xy_array, out outids_sa),
                VisioObjectCategory.Shape => this.Shape.DropManyU(masters_obj_array, xy_array, out outids_sa),
                _ => throw new System.ArgumentException(_unhandled_category_exc_msg)
            };

            short[] outids = (short[]) outids_sa;
            return outids;
        }

        public IVisio.Shape Drop(
            IVisio.Master master,
            Point point)
        {
            if (master == null)
            {
                throw new System.ArgumentNullException(nameof(master));
            }

            var shape = this.Category switch
            {
                VisioObjectCategory.Master => this.Master.Drop(master, point.X, point.Y),
                VisioObjectCategory.Page => this.Page.Drop(master, point.X, point.Y),
                VisioObjectCategory.Shape => this.Shape.Drop(master, point.X, point.Y),
                _ => throw new System.ArgumentException(_unhandled_category_exc_msg)
            };

            return shape;

        }

        public int SetFormulas(ShapeSheet.Streams.StreamArray stream, object[] formulas, short flags)
        {
            Internal.TempHelper.ValidateStreamLengthFormulas(stream, formulas);

            var val = this.Category switch
            {
                VisioObjectCategory.Master => this.Master.SetFormulas(stream, formulas, flags),
                VisioObjectCategory.Page => this.Page.SetFormulas(stream, formulas, flags),
                VisioObjectCategory.Shape => this.Shape.SetFormulas(stream, formulas, flags),
                _ => throw new System.ArgumentException(_unhandled_category_exc_msg)
            };

            return val;
        }



        public int SetResults(ShapeSheet.Streams.StreamArray stream, object[] unitcodes, object[] results, short flags)
        {
            Internal.TempHelper.ValidateStreamLengthResults(stream, results);

            var val = this.Category switch
            {
                VisioObjectCategory.Master => this.Master.SetResults(stream, unitcodes, results, flags),
                VisioObjectCategory.Page => this.Page.SetResults(stream, unitcodes, results, flags),
                VisioObjectCategory.Shape => this.Shape.SetResults(stream, unitcodes, results, flags),
                _ => throw new System.ArgumentException(_unhandled_category_exc_msg)
            };

            return val;
        }

        public TResult[] GetResults<TResult>(ShapeSheet.Streams.StreamArray stream, object[] unitcodes)
        {
            Internal.TempHelper._enforce_valid_result_type(typeof(TResult));

            var val = this.Category switch
            {
                VisioObjectCategory.Master => this.Master.GetResults<TResult>(stream, unitcodes),
                VisioObjectCategory.Page => this.Page.GetResults<TResult>(stream, unitcodes),
                VisioObjectCategory.Shape => this.Shape.GetResults<TResult>(stream, unitcodes),
                _ => throw new System.ArgumentException(_unhandled_category_exc_msg)
            };

            return val;
        }

        public string[] GetFormulasU(ShapeSheet.Streams.StreamArray stream)
        {
            var val = this.Category switch
            {
                VisioObjectCategory.Master => this.Master.GetFormulasU(stream),
                VisioObjectCategory.Page => this.Page.GetFormulasU(stream),
                VisioObjectCategory.Shape => this.Shape.GetFormulasU(stream),
                _ => throw new System.ArgumentException(_unhandled_category_exc_msg)
            };

            return val;
        }

        internal static T[] system_array_to_typed_array<T>(System.Array results_sa)
        {
            var results = new T[results_sa.Length];
            results_sa.CopyTo(results, 0);
            return results;
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
                throw new Exceptions.InternalAssertionException(msg);
            }
            return flags;
        }

    }
}
