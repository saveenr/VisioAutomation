using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VTest.Models
{
    [MUT.TestClass]
    public class XmlModelDrawModelTests : Framework.VTest
    {
        [MUT.TestMethod]
        public void DrawXmlModel_TopNodeIsLabelledWithTheDocumentElementName()
        {
            var xml = new System.Xml.XmlDocument();
            xml.LoadXml("<root a=\"1\"><child1><leaf/></child1><child2>text</child2></root>");
            var model = new VisioAutomation.Models.Data.XmlModel();
            model.XmlDocument = xml;

            var client = this.GetScriptingClient();
            client.Document.NewDocument();
            client.Model.DrawXmlModel(VisioScripting.TargetPage.Auto, model);

            var labels = this.GetVisioApplication().ActivePage.Shapes.Cast<IVisio.Shape>()
                .Select(s => s.Text)
                .Where(t => t != string.Empty)
                .OrderBy(t => t)
                .ToList();

            MUT.CollectionAssert.AreEqual(new[] { "child1", "child2", "leaf", "root" }, labels);

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void DrawXmlModel_IsUndoneByASingleUndo()
        {
            var xml = new System.Xml.XmlDocument();
            xml.LoadXml("<root><child1><leaf/></child1><child2/></root>");
            var model = new VisioAutomation.Models.Data.XmlModel();
            model.XmlDocument = xml;

            var client = this.GetScriptingClient();
            client.Document.NewDocument();
            var page = this.GetVisioApplication().ActivePage;
            int shapes_before = page.Shapes.Count;

            client.Model.DrawXmlModel(VisioScripting.TargetPage.Auto, model);
            MUT.Assert.IsTrue(page.Shapes.Count > shapes_before, "the model should have drawn shapes");

            client.Undo.UndoLastAction();

            MUT.Assert.AreEqual(shapes_before, page.Shapes.Count, "one Undo should remove everything the draw added");

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void DrawXmlModel_DocumentWithoutADocumentElement_ThrowsArgumentException()
        {
            var model = new VisioAutomation.Models.Data.XmlModel();
            model.XmlDocument = new System.Xml.XmlDocument();

            var client = this.GetScriptingClient();
            client.Document.NewDocument();

            MUT.Assert.ThrowsExactly<System.ArgumentException>(
                () => client.Model.DrawXmlModel(VisioScripting.TargetPage.Auto, model));

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }
    }
}
