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

    }
}
