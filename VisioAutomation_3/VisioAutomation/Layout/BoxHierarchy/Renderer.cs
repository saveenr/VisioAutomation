using VA=VisioAutomation;


namespace VisioAutomation.Layout.BoxHierarchy
{
    internal class Renderer<T>
    {
        public RenderOptions<T> RenderOptions { get; set; }

        public Renderer()
        {
            this.RenderOptions = new RenderOptions<T>();
        }

        public void Render(BoxHierarchyLayout<T> layout)
        {
            this.Render(layout.Root);
        }

        public void Render(Node<T> node)
        {
            if (node == null)
            {
                throw new System.ArgumentNullException("node");
            }

            if (this.RenderOptions == null)
            {
                throw new System.ArgumentException("renderoptions is null");
            }

            if (this.RenderOptions.RenderAction == null)
            {
                throw new System.ArgumentException("renderoptions contains a null function");
            }

            this.RenderOptions.RenderAction(node, node.Rectangle);

            foreach (var cur_el in node.Children)
            {
                Render(cur_el);
            }
        }
    }
}