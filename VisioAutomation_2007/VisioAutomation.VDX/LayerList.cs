namespace VisioAutomation.VDX
{
    public class LayerList : NamedNodeList<Elements.Layer>
    {
        public LayerList() :
            base(layer => layer.Name)
        {

        }
    }
}