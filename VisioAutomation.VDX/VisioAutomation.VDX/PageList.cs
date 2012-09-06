namespace VisioAutomation.VDX
{
    public class PageList : NamedNodeList<Elements.Page>
    {
        private Elements.Drawing drawing_el;
        public PageList(Elements.Drawing drawing_el) :
            base(page => page.Name)
        {
            this.drawing_el = drawing_el;
        }

        public override void Add(Elements.Page page)
        {
            page.Drawing = this.drawing_el;
            base.Add(page);
        }
    }
}