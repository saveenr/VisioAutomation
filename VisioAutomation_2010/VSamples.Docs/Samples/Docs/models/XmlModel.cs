using IVisio = Microsoft.Office.Interop.Visio;
using VAMODELS = VisioAutomation.Models;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page models/xml-model.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ModelsXmlModelSamples
    {
        public static void Sample_01(VisioScripting.Client client)
        {
            // <sample id="models/xml-model#01" doc="models/xml-model.md" line="19">
            var xml = new System.Xml.XmlDocument();
            xml.LoadXml("<root><child1><leaf/></child1><child2/></root>");

            var model = new VAMODELS.Data.XmlModel();
            model.XmlDocument = xml;

            client.Model.DrawXmlModel(VisioScripting.TargetPage.Auto, model);
            // </sample>
        }
    }
}
