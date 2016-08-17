using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet
{
    public struct ShapeSheetSurface
    {
        public SurfaceTarget Target { get; private set; }

        public ShapeSheetSurface(SurfaceTarget target)
        {
            this.Target = target;
        }

        public ShapeSheetSurface(IVisio.Page page)
        {
            this.Target = new SurfaceTarget(page);
        }

        public ShapeSheetSurface(IVisio.Master master)
        {
            this.Target = new SurfaceTarget(master);
        }

        public ShapeSheetSurface(IVisio.Shape shape)
        {
            this.Target = new SurfaceTarget(shape);
        }

        public int SetFormulas(short[] stream, object[] formulas, short flags)
        {
            int c;
            if (this.Target.Shape != null)
            {
                c = this.Target.Shape.SetFormulas(stream, formulas, flags);
            }
            else if (this.Target.Master != null)
            {
                c = this.Target.Master.SetFormulas(stream, formulas, flags);
            }
            else if (this.Target.Page != null)
            {
                c = this.Target.Page.SetFormulas(stream, formulas, flags);
            }
            else
            {
                throw new System.ArgumentException("Unhandled Target");
            }

            return c;
        }

        public int SetResults(short[] stream, object[] unitcodes, object[] results, short flags)
        {
            int c;
            if (this.Target.Shape != null)
            {
                c = this.Target.Shape.SetResults(stream, unitcodes, results, flags);
            }
            else if (this.Target.Master != null)
            {
                c = this.Target.Master.SetResults(stream, unitcodes, results, flags);
            }
            else if (this.Target.Page != null)
            {
                c = this.Target.Page.SetResults(stream, unitcodes, results, flags);
            }
            else
            {
                throw new System.ArgumentException("Unhandled Target");
            }

            return c;
        }

    }
}