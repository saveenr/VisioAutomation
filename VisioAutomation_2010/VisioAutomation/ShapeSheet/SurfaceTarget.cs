using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet
{
    public struct SurfaceTarget
    {
        public IVisio.Page Page;
        public IVisio.Master Master;
        public IVisio.Shape Shape;

        public SurfaceTarget(IVisio.Page page)
        {
            this.Page = page;
            this.Master = null;
            this.Shape = null;
        }

        public SurfaceTarget(IVisio.Master master)
        {
            this.Page = null;
            this.Master = master;
            this.Shape = null;
        }

        public SurfaceTarget(IVisio.Shape shape)
        {
            this.Page = null;
            this.Master = null;
            this.Shape = shape;
        }

        public SurfaceTarget(IVisio.Page page, IVisio.Master master, IVisio.Shape shape)
        {
            this.Page = page;
            this.Master = master;
            this.Shape = shape;
        }
    }
}