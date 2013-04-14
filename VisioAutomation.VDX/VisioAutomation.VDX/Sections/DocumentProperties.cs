using VisioAutomation.VDX.Internal.Extensions;
using VisioAutomation.VDX.Internal;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections
{
    public class DocumentProperties
    {
        public string Creator = string.Empty;
        public string Company = string.Empty;
        public int? BuildNumberCreated = 805312791;
        public int? BuildNumberEdited = 805312791;
        // TODO: Add support for PreviewPicture
        // TODO: Add support for CustomProps
        public System.DateTimeOffset? TimeCreated;
        public System.DateTimeOffset? TimeSaved;
        public System.DateTimeOffset? TimeEdited;
        public System.DateTimeOffset? TimePrinted;

        public SXL.XElement ToXml()
        {

            var ns = Constants.VisioXmlNamespace2003;

            var el = XMLUtil.CreateVisioSchema2003Element("DocumentProperties");
            el.SetElementValue(ns + "Creator", this.Creator);
            el.SetElementValue(ns + "Company", this.Company);
            el.SetElementValueConditional(ns + "BuildNumberCreated", this.BuildNumberCreated);
            el.SetElementValueConditional(ns + "BuildNumberEdited", this.BuildNumberEdited);
            el.SetElementValueConditionalDateTime(ns+"TimeCreated", this.TimeCreated);
            el.SetElementValueConditionalDateTime(ns + "TimeSaved", this.TimeSaved);
            el.SetElementValueConditionalDateTime(ns + "TimeEdited", this.TimeEdited);
            el.SetElementValueConditionalDateTime(ns + "TimePrinted", this.TimePrinted);

            return el;
        }
    }
}