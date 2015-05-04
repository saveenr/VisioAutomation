using VisioAutomation.VDX.Internal;
using SXL=System.Xml.Linq;

namespace VisioAutomation.VDX.ShapeSheet
{
    public abstract class CellBase
    {
        public string Formula { get; set; }
        public bool InheritFormula { get; set; }
        public CellUnit Unit { get; set; }

        public abstract string GetResultString();
        public abstract bool HasResult { get; }

        protected CellBase()
        {
            this.Unit = CellUnit.None;
        }

        protected CellBase(CellUnit unit)
        {
            this.Unit = unit;
        }

        private SXL.XElement _ToXml(string elname)
        {
            if (this.Formula != null || this.HasResult)
            {
                var el = new SXL.XElement(elname);

                if (this.HasResult)
                {
                    el.Value = this.GetResultString();
                }

                if (this.Formula != null)
                {
                    el.SetAttributeValue("F", this.Formula);
                }

                if (this.Unit == CellUnit.Inch)
                {
                    el.SetAttributeValue("Unit", "IN");
                }
                else if (this.Unit == CellUnit.Point)
                {
                    el.SetAttributeValue("Unit", "PT");
                }

                return el;
            }
            return null;
        }

        public SXL.XElement ToXml(string elname)
        {
            if (elname.StartsWith("{"))
            {
                throw new System.ArgumentException("elname");
            }

            var fullname = string.Format("{0}{1}",Constants.VisioXmlNamespace2003, elname);
            return this._ToXml(fullname);                
        }

        public SXL.XElement ToXml2006(string elname)
        {
            if (elname.StartsWith("{"))
            {
                throw new System.ArgumentException("elname");
            }
            var fullname = string.Format("{0}{1}", Constants.VisioXmlNamespace2006, elname);
            return this._ToXml(fullname);
        }

    }
}