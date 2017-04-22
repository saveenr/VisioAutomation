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

    }
}