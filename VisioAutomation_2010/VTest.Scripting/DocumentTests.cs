using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using VisioAutomation.Extensions;

namespace VTest.Scripting
{
    [MUT.TestClass]
    public class DocumentTests : Framework.VTest
    {
        [MUT.TestMethod]
        public void ActivateDocument_AmongMultipleOpenDocs_UpdatesActiveDocument()
        {
            var client = this.GetScriptingClient();
            var app = client.Application.GetApplication();
            var doc1 = client.Document.NewDocument();
            var doc2 = client.Document.NewDocument();
            var doc3 = client.Document.NewDocument();

            client.Document.ActivateDocument(doc1);
            MUT.Assert.AreEqual(doc1, app.ActiveDocument);
            client.Document.ActivateDocument(doc2);
            MUT.Assert.AreEqual(doc2, app.ActiveDocument);
            client.Document.ActivateDocument(doc3);
            MUT.Assert.AreEqual(doc3, app.ActiveDocument);
            client.Document.ActivateDocument(doc1);
            MUT.Assert.AreEqual(doc1, app.ActiveDocument);

            doc1.Close(true);
            doc2.Close(true);
            doc3.Close(true);
        }

        // -- NewDocumentFromTemplate (#229) ------------------------------------------------------
        // Visio 2013 and later ship XML templates (.vstx); older versions ship binary ones (.vst).

        private static string flowchart_template(VisioScripting.Client client)
        {
            return client.Application.ApplicationVersion.Major >= 15 ? "basflo_u.vstx" : "basflo_u.vst";
        }

        private static int count_empty_stencils(Microsoft.Office.Interop.Visio.Application app)
        {
            // A document that is a stencil and holds no masters is the stray window the old code left behind
            return app.Documents.Cast<Microsoft.Office.Interop.Visio.Document>().Count(
                d => d.Type == Microsoft.Office.Interop.Visio.VisDocumentTypes.visTypeStencil && d.Masters.Count == 0);
        }

        [MUT.TestMethod]
        public void NewDocumentFromTemplate_RealTemplate_CreatesTheDocumentFromTheTemplate()
        {
            // Regression: the template was opened as a docked stencil next to a blank drawing, so the
            // drawing was not based on it and an empty stencil was left open.
            var client = this.GetScriptingClient();
            var app = client.Application.GetApplication();
            int empty_stencils_before = count_empty_stencils(app);

            var doc = client.Document.NewDocumentFromTemplate(flowchart_template(client));

            MUT.Assert.AreEqual(Microsoft.Office.Interop.Visio.VisDocumentTypes.visTypeDrawing, doc.Type);
            MUT.Assert.IsTrue(doc.Template.ToLowerInvariant().Contains("basflo_u"), "the drawing should be based on the template: '" + doc.Template + "'");
            MUT.Assert.AreEqual(empty_stencils_before, count_empty_stencils(app), "no empty stencil should be left open");

            doc.Close(true);
        }

        [MUT.TestMethod]
        public void NewDocumentFromTemplate_NullTemplate_CreatesABlankDrawing()
        {
            var client = this.GetScriptingClient();
            var app = client.Application.GetApplication();
            int docs_before = app.Documents.Count;

            var doc = client.Document.NewDocumentFromTemplate(null);

            MUT.Assert.AreEqual(Microsoft.Office.Interop.Visio.VisDocumentTypes.visTypeDrawing, doc.Type);
            MUT.Assert.AreEqual(string.Empty, doc.Template);
            MUT.Assert.AreEqual(docs_before + 1, app.Documents.Count, "exactly one document should be added");

            doc.Close(true);
        }

        [MUT.TestMethod]
        public void NewDocumentFromTemplate_EmptyOrWhitespaceTemplate_CreatesABlankDrawing()
        {
            var client = this.GetScriptingClient();
            var app = client.Application.GetApplication();

            foreach (string template in new[] { "", "   " })
            {
                int docs_before = app.Documents.Count;
                var doc = client.Document.NewDocumentFromTemplate(template);

                MUT.Assert.AreEqual(Microsoft.Office.Interop.Visio.VisDocumentTypes.visTypeDrawing, doc.Type);
                MUT.Assert.AreEqual(docs_before + 1, app.Documents.Count, "exactly one document should be added for '" + template + "'");

                doc.Close(true);
            }
        }

        [MUT.TestMethod]
        public void NewDocumentFromTemplate_StencilFile_ThrowsAndCreatesNoDocument()
        {
            // A stencil is not a template. The old code threw an opaque COM error after leaving a blank drawing open.
            var client = this.GetScriptingClient();
            var app = client.Application.GetApplication();

            foreach (string stencil in new[] { "basic_u.vss", "BASIC_U.VSSX", @"C:\Stencils\mine.vssx" })
            {
                int docs_before = app.Documents.Count;

                var ex = MUT.Assert.ThrowsExactly<System.ArgumentException>(
                    () => client.Document.NewDocumentFromTemplate(stencil));

                MUT.StringAssert.Contains(ex.Message.ToLowerInvariant(), "stencil");
                MUT.Assert.AreEqual(docs_before, app.Documents.Count, "no document should be created for '" + stencil + "'");
            }
        }

        [MUT.TestMethod]
        public void NewDocumentFromTemplate_MissingTemplate_ThrowsAndCreatesNoDocument()
        {
            var client = this.GetScriptingClient();
            var app = client.Application.GetApplication();
            int docs_before = app.Documents.Count;

            bool threw = false;
            try
            {
                client.Document.NewDocumentFromTemplate("no_such_template_for_vtest.vstx");
            }
            catch (System.Exception)
            {
                threw = true;
            }

            MUT.Assert.IsTrue(threw, "a template that does not exist should throw");
            MUT.Assert.AreEqual(docs_before, app.Documents.Count, "no stray blank drawing should be left open");
        }
    }
}