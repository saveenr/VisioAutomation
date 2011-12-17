using System;
using SXL=System.Xml.Linq;

namespace VisioAutomation.VDX.Internal
{
    public static class XMLUtil
    {
        public static SXL.XElement CreateVisioSchema2003Element(string name)
        {
            string fullname = String.Format("{0}{1}",Constants.VisioXmlNamespace2003,name);
            var el = new SXL.XElement(fullname);
            return el;
        }

        public static SXL.XElement CreateVisioSchema2006Element(string name)
        {
            string fullname = String.Format("{0}{1}", Constants.VisioXmlNamespace2006, name);
            var el = new SXL.XElement(fullname);
            return el;
        }

    }

}