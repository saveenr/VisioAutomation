using System.Diagnostics.Eventing.Reader;
using IVisio = Microsoft.Office.Interop.Visio;


namespace VisioAutomation.Internal
{
    internal readonly struct VisioObjectTarget
    {
        public readonly IVisio.Page Page;
        public readonly IVisio.Master Master;
        public readonly IVisio.Shape Shape;
        public readonly VisioObjectCategory Category;
        private static readonly string _unhandled_category_exc_msg = string.Format("Unhandled {0}",nameof(VisioObjectCategory));

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

        public void DispatchAction(
            System.Action<IVisio.Shape> fshape,
            System.Action<IVisio.Master> fmaster,
            System.Action<IVisio.Page> fpage)
        {
            if (this.Category == VisioObjectCategory.Shape)
            {
                fshape(this.Shape);
            }
            else if (this.Category == VisioObjectCategory.Master)
            {
                fmaster(this.Master);
            }
            else if (this.Category == VisioObjectCategory.Page)
            {
                fpage(this.Page);
            }
            else
            {
                throw new System.ArgumentException(_unhandled_category_exc_msg);
            }
        }

        public void DispatchAction<P1>(
            System.Action<IVisio.Shape,P1> fshape,
            System.Action<IVisio.Master, P1> fmaster,
            System.Action<IVisio.Page, P1> fpage,
            P1 p1)
        {
            if (this.Category == VisioObjectCategory.Shape)
            {
                fshape(this.Shape,p1);
            }
            else if (this.Category == VisioObjectCategory.Master)
            {
                fmaster(this.Master,p1);
            }
            else if (this.Category == VisioObjectCategory.Page)
            {
                fpage(this.Page,p1);
            }
            else
            {
                throw new System.ArgumentException(_unhandled_category_exc_msg);
            }
        }

        public void DispatchAction<P1,P2>(
            System.Action<IVisio.Shape, P1, P2> fshape,
            System.Action<IVisio.Master, P1, P2> fmaster,
            System.Action<IVisio.Page, P1, P2> fpage,
            P1 p1,
            P2 p2)
        {
            if (this.Category == VisioObjectCategory.Shape)
            {
                fshape(this.Shape, p1, p2);
            }
            else if (this.Category == VisioObjectCategory.Master)
            {
                fmaster(this.Master, p1, p2);
            }
            else if (this.Category == VisioObjectCategory.Page)
            {
                fpage(this.Page, p1, p2);
            }
            else
            {
                throw new System.ArgumentException(_unhandled_category_exc_msg);
            }
        }



        public T DispatchFunction<T>(
            System.Func<IVisio.Shape, T> fshape,
            System.Func<IVisio.Master, T> fmaster, 
            System.Func<IVisio.Page, T> fpage)
        {
            T res = this.Category switch
            {
                VisioObjectCategory.Shape => fshape(this.Shape),
                VisioObjectCategory.Master => fmaster(this.Master),
                VisioObjectCategory.Page => fpage(this.Page),
                _ => throw new System.ArgumentException(_unhandled_category_exc_msg)
            };
            return res;
        }

        public T DispatchFunction<P1,T>(
            System.Func<IVisio.Shape, P1, T> fshape,
            System.Func<IVisio.Master, P1, T> fmaster,
            System.Func<IVisio.Page, P1, T> fpage,
            P1 p1)
        {
            T res = this.Category switch
            {
                VisioObjectCategory.Shape => fshape(this.Shape, p1),
                VisioObjectCategory.Master => fmaster(this.Master, p1),
                VisioObjectCategory.Page => fpage(this.Page, p1),
                _ => throw new System.ArgumentException(_unhandled_category_exc_msg)
            };
            return res;
        }


        public T DispatchFunction<P1, P2, T>(
            System.Func<IVisio.Shape, P1, P2, T> fshape,
            System.Func<IVisio.Master, P1, P2, T> fmaster,
            System.Func<IVisio.Page, P1, P2, T> fpage,
            P1 p1,
            P2 p2)
        {
            T res = this.Category switch
            {
                VisioObjectCategory.Shape => fshape(this.Shape, p1,p2),
                VisioObjectCategory.Master => fmaster(this.Master, p1,p2),
                VisioObjectCategory.Page => fpage(this.Page, p1,p2),
                _ => throw new System.ArgumentException(_unhandled_category_exc_msg)
            };
            return res;
        }

        public T DispatchFunction<P1, P2, P3, T>(
            System.Func<IVisio.Shape, P1, P2, P3, T> fshape,
            System.Func<IVisio.Master, P1, P2, P3, T> fmaster,
            System.Func<IVisio.Page, P1, P2, P3, T> fpage,
            P1 p1,
            P2 p2,
            P3 p3)
        {
            T res = this.Category switch
            {
                VisioObjectCategory.Shape => fshape(this.Shape, p1, p2, p3),
                VisioObjectCategory.Master => fmaster(this.Master, p1, p2, p3),
                VisioObjectCategory.Page => fpage(this.Page, p1, p2, p3),
                _ => throw new System.ArgumentException(_unhandled_category_exc_msg)
            };
            return res;
        }

        public T DispatchFunction<P1, P2, P3, P4, T>(
            System.Func<IVisio.Shape, P1, P2, P3, P4, T> fshape,
            System.Func<IVisio.Master, P1, P2, P3, P4, T> fmaster,
            System.Func<IVisio.Page, P1, P2, P3, P4, T> fpage,
            P1 p1,
            P2 p2,
            P3 p3,
            P4 p4)
        {
            T res = this.Category switch
            {
                VisioObjectCategory.Shape => fshape(this.Shape, p1, p2, p3, p4),
                VisioObjectCategory.Master => fmaster(this.Master, p1, p2, p3, p4),
                VisioObjectCategory.Page => fpage(this.Page, p1, p2, p3, p4),
                _ => throw new System.ArgumentException(_unhandled_category_exc_msg)
            };
            return res;
        }
    }
}
