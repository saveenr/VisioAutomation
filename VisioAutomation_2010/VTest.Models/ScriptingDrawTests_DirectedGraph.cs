using System.IO;
using VTest.Framework;
using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;
using SXL = System.Xml.Linq;

namespace VTest.Models
{
    public partial class ScriptingDrawTests : Framework.VTest
    {

        [MUT.TestMethod]
        [MUT.DeploymentItem(@"datafiles\directed_graph_1.xml", "datafiles")]
        public void Scripting_Draw_DirectedGraph1()
        {
            // Load the graph
            string xml = this.get_datafile_content(@"datafiles\directed_graph_1.xml");

            // Draw the graph
            var client = this.GetScriptingClient();
            this.draw_directed_graph(client, xml);

            // Cleanup
            string output_filename = VTestGlobals.VTestHelper.GetOutputFilename(nameof(Scripting_Draw_DirectedGraph1),".vsd");

            client.Document.SaveDocumentAs(VisioScripting.TargetDocument.Auto, output_filename);
            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        [MUT.DeploymentItem(@"datafiles\directed_graph_2.xml", "datafiles")]
        public void Scripting_Draw_DirectedGraph2()
        {
            // Load the graph
            string xml = this.get_datafile_content(@"datafiles\directed_graph_2.xml");

            // Draw the graph
            var client = this.GetScriptingClient();
            this.draw_directed_graph(client, xml);

            // Cleanup
            string output_filename = VTestGlobals.VTestHelper.GetOutputFilename(nameof(Scripting_Draw_DirectedGraph2),".vsd");
            client.Document.SaveDocumentAs(VisioScripting.TargetDocument.Auto, output_filename);
            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        [MUT.DeploymentItem(@"datafiles\directed_graph_3.xml", "datafiles")]
        public void Scripting_Draw_DirectedGraph3()
        {
            // Load the graph
            string xml = this.get_datafile_content(@"datafiles\directed_graph_3.xml");

            // Draw the graph
            var client = this.GetScriptingClient();
            this.draw_directed_graph(client, xml);

            // Cleanup
            string output_filename = VTestGlobals.VTestHelper.GetOutputFilename(nameof(Scripting_Draw_DirectedGraph3),".vsd");

            client.Document.SaveDocumentAs(VisioScripting.TargetDocument.Auto, output_filename);
            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        [MUT.DeploymentItem(@"datafiles\directed_graph_4.xml", "datafiles")]
        public void Scripting_Draw_DirectedGraph4()
        {
            // Load the graph
            string xml = this.get_datafile_content(@"datafiles\directed_graph_4.xml");

            // Draw the graph
            var client = this.GetScriptingClient();
            this.draw_directed_graph(client, xml);

            // Cleanup
            string output_filename = VTestGlobals.VTestHelper.GetOutputFilename(nameof(Scripting_Draw_DirectedGraph4),".vsd");

            client.Document.SaveDocumentAs(VisioScripting.TargetDocument.Auto, output_filename);
            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        public string get_datafile_content(string name)
        {
            string inputfilename = this._get_test_results_out_path( name );

            if (!File.Exists(inputfilename))
            {
                MUT.Assert.Fail("Could not locate " + inputfilename);
            }
            string text = File.ReadAllText(inputfilename);
            return text;
        }

        private void draw_directed_graph(VisioScripting.Client client, string dg_text)
        {
            var dg_xml = SXL.XDocument.Parse(dg_text);
            var dgdoc = VisioScripting.Builders.DirectedGraphDocumentLoader.LoadFromXml(client, dg_xml);

            // TODO: Investigate if this this special case for Visio 2013 can be removed
            // this is a temporary fix to handle the fact that server_u.vss in Visio 2013 doesn't result in server_u.vssx 
            // getting automatically loaded

            var version = client.Application.ApplicationVersion;
            if (version.Major >= 15)
            {
                foreach (var drawing in dgdoc.Layouts)
                {
                    foreach (var shape in drawing.Nodes)
                    {
                        if (shape.StencilName == "server_u.vss")
                        {
                            shape.StencilName = "server_u.vssx";
                        }
                    }
                }
            }

            var dgstyling = new VA.Models.Layouts.DirectedGraph.DirectedGraphStyling();

            client.Model.DrawDirectedGraphDocument(dgdoc,dgstyling);
        }
        
    }
}