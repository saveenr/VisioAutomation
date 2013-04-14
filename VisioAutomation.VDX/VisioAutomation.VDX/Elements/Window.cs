using VisioAutomation.VDX.Internal;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Elements
{
    public class Window
    {
        private static IDGenerator idgen = new IDGenerator(0);

        private int _id;
        public int? Width { get; set; }
        public int? Height { get; set; }

        protected Window()
        {
            this._id = idgen.GetNextID();
        }

        public int ID
        {
            get { return _id; }
        }

        public virtual void AddToElement(SXL.XElement parent)
        {
            throw new System.Exception();
        }
    }
}