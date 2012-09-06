namespace VisioAutomation.VDX
{
    public class FaceList : NamedNodeList<Elements.Face>
    {
        public FaceList() :
            base(face => face.Name)
        {

        }
    }
}