using VisioAutomation.VDX.Internal;

namespace VisioAutomation.VDX.ShapeSheet
{
    public abstract class CellBase
    {
        public string Formula;
        public bool InheritFormula;
        public CellUnit Unit;

        public CellBase()
        {
            this.Unit = CellUnit.None;
        }

        public CellBase(CellUnit unit)
        {
            this.Unit = unit;
        }

        public abstract string GetResultString();
        public abstract bool HasResult { get; }

        public System.Xml.Linq.XElement _ToXml(string elname)
        {
            if (this.Formula != null || this.HasResult)
            {
                var el = new System.Xml.Linq.XElement(elname);
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


        public System.Xml.Linq.XElement ToXml(string elname)
        {
            if (elname.StartsWith("{"))
            {
                throw new System.ArgumentException();
            }
            var fullname = string.Format("{0}{1}",Constants.VisioXmlNamespace2003, elname);
            return this._ToXml(fullname);                
        }

        public System.Xml.Linq.XElement ToXml2006(string elname)
        {
            if (elname.StartsWith("{"))
            {
                throw new System.ArgumentException();
            }
            var fullname = string.Format("{0}{1}", Constants.VisioXmlNamespace2006, elname);
            return this._ToXml(fullname);
        }

    }
}