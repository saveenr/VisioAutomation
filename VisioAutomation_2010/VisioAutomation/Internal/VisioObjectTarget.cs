using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Core;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Internal
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
            var outids = this.Category switch
            {
                VisioObjectCategory.Master => this.Master.DropManyU(masters, points),
                VisioObjectCategory.Page => this.Page.DropManyU(masters, points),
                VisioObjectCategory.Shape => this.Shape.DropManyU(masters, points),
                _ => throw new System.ArgumentException(_unhandled_category_exc_msg)
            };

            return outids;
        }

        public int SetFormulas(ShapeSheet.Streams.StreamArray stream, object[] formulas, short flags)
        {
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


    }
}
