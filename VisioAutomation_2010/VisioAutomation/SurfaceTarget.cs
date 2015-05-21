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
            if (master== null)
            {
                throw new System.ArgumentNullException(nameof(master));
            }

            this.Page = null;
            this.Master = master;
            this.Shape = null;
        }

        public SurfaceTarget(IVisio.Shape shape)
        {
            if (shape== null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            this.Page = null;
            this.Master = null;
            this.Shape = shape;
        }
    }
}