using System.Collections.Generic;
using VA=VisioAutomation;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Elements
{
    public class Hyperlink
    {
        public string Description;
        public string Address;
        public string SubAddress;

        public Hyperlink(string description, string address, string subaddress)
        {
            this.Address = address;
            this.Description = description;
            this.SubAddress = subaddress;
        }

        public Hyperlink(string description, string address)
        {
            this.Address = address;
            this.Description = description;
            this.SubAddress = null;
        }
    }
}