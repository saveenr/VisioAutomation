using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Linq;

namespace TestVisioAutomation
{
    [TestClass]
    public class AutoLayoutTests : VisioAutomationTest
    {
        [TestMethod]
        public void RenderDirectedGraphWithBezierConnectors()
        {
            var directed_graph_drawing = new VA.Layout.Models.DirectedGraph.Drawing();

            var flowchart_stencil = "basflo_u.vss";
            var server_stencil = "server_u.vss";

            var n0 = directed_graph_drawing.AddShape("n0", "Untitled Node", flowchart_stencil, "Decision");
            n0.Size = new VA.Drawing.Size(3, 2);
            var n1 = directed_graph_drawing.AddShape("n1", "", flowchart_stencil, "Decision");
            var n2 = directed_graph_drawing.AddShape("n2", "MailServer", server_stencil, "Server");
            var n3 = directed_graph_drawing.AddShape("n3", null, flowchart_stencil, "Data");
            var n4 = directed_graph_drawing.AddShape("n4", "Alone", flowchart_stencil, "Data");

            var c0 = directed_graph_drawing.Connect("c0", n0, n1, null, VA.Connections.ConnectorType.Curved);
            var c1 = directed_graph_drawing.Connect("c1", n1, n2, "YES", VA.Connections.ConnectorType.RightAngle);
            var c3 = directed_graph_drawing.Connect("c2", n1, n0, "NO", VA.Connections.ConnectorType.Curved);
            var c4 = directed_graph_drawing.Connect("c3", n0, n2, null, VA.Connections.ConnectorType.Straight);
            var c5 = directed_graph_drawing.Connect("c4", n2, n3, null, VA.Connections.ConnectorType.Curved);
            var c6 = directed_graph_drawing.Connect("c5", n3, n0, null, VA.Connections.ConnectorType.Curved);


            var options = new VA.Layout.MSAGL.LayoutOptions();
            options.UseDynamicConnectors = false;
            var visapp = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = visapp.ActivePage;

            VA.Layout.MSAGL.MSAGLRenderer.Render(page, directed_graph_drawing, options);

            Assert.IsNotNull(n0.VisioShape);
            Assert.IsNotNull(n1.VisioShape);
            Assert.IsNotNull(n2.VisioShape);
            Assert.IsNotNull(n3.VisioShape);
            Assert.IsNotNull(n4.VisioShape);
            Assert.IsNotNull(c0.VisioShape);
            Assert.IsNotNull(c1.VisioShape);
            Assert.IsNotNull(c3.VisioShape);
            Assert.IsNotNull(c4.VisioShape);
            Assert.IsNotNull(c5.VisioShape);
            Assert.IsNotNull(c6.VisioShape);



            // Decision 
            var test_shape_0 = page.Shapes["Decision"];
            //Assert.AreEqual("11.523988095238 in.", test_shape_0.CellsU["PinX"].Formula);
            //Assert.AreEqual("11.931904761905 in.", test_shape_0.CellsU["PinY"].Formula);
            Assert.AreEqual("Width*0.5", test_shape_0.CellsU["LocPinX"].Formula);
            Assert.AreEqual("Height*0.5", test_shape_0.CellsU["LocPinY"].Formula);
            Assert.AreEqual("3 in.", test_shape_0.CellsU["Width"].Formula);
            Assert.AreEqual("2 in.", test_shape_0.CellsU["Height"].Formula);
            Assert.AreEqual("0 deg.", test_shape_0.CellsU["Angle"].Formula);
            Assert.AreEqual("Untitled Node", test_shape_0.Text);

            // Decision.2 
            var test_shape_1 = page.Shapes["Decision.2"];
            //Assert.AreEqual("11.523988095238 in.", test_shape_1.CellsU["PinX"].Formula);
            //Assert.AreEqual("7.9140476190476 in.", test_shape_1.CellsU["PinY"].Formula);
            Assert.AreEqual("Width*0.5", test_shape_1.CellsU["LocPinX"].Formula);
            Assert.AreEqual("Height*0.5", test_shape_1.CellsU["LocPinY"].Formula);
            Assert.AreEqual("1 in.", test_shape_1.CellsU["Width"].Formula);
            Assert.AreEqual("0.75 in.", test_shape_1.CellsU["Height"].Formula);
            Assert.AreEqual("0 deg.", test_shape_1.CellsU["Angle"].Formula);
            Assert.AreEqual("", test_shape_1.Text);

            // Server 
            var test_shape_2 = page.Shapes["Server"];
            //Assert.AreEqual("6.3097023809524 in.", test_shape_2.CellsU["PinX"].Formula);
            //Assert.AreEqual("4.3961904761905 in.", test_shape_2.CellsU["PinY"].Formula);
            Assert.AreEqual("Width*0.5", test_shape_2.CellsU["LocPinX"].Formula);
            Assert.AreEqual("Height*0.5", test_shape_2.CellsU["LocPinY"].Formula);
            Assert.AreEqual("0.7 in.", test_shape_2.CellsU["Width"].Formula);
            Assert.AreEqual("1 in.", test_shape_2.CellsU["Height"].Formula);
            Assert.AreEqual("0 deg.", test_shape_2.CellsU["Angle"].Formula);
            Assert.AreEqual("MailServer", test_shape_2.Text);

            // Data 
            var test_shape_3 = page.Shapes["Data"];
            //Assert.AreEqual("1.0954166666667 in.", test_shape_3.CellsU["PinX"].Formula);
            //Assert.AreEqual("0.87833333333334 in.", test_shape_3.CellsU["PinY"].Formula);
            Assert.AreEqual("Width*0.5", test_shape_3.CellsU["LocPinX"].Formula);
            Assert.AreEqual("Height*0.5", test_shape_3.CellsU["LocPinY"].Formula);
            Assert.AreEqual("1 in.", test_shape_3.CellsU["Width"].Formula);
            Assert.AreEqual("0.75 in.", test_shape_3.CellsU["Height"].Formula);
            Assert.AreEqual("0 deg.", test_shape_3.CellsU["Angle"].Formula);
            Assert.AreEqual("", test_shape_3.Text);

            // Data.19 
            var test_shape_4 = page.Shapes["Data.19"];
            //Assert.AreEqual("4.7382738095238 in.", test_shape_4.CellsU["PinX"].Formula);
            //Assert.AreEqual("0.87833333333334 in.", test_shape_4.CellsU["PinY"].Formula);
            Assert.AreEqual("Width*0.5", test_shape_4.CellsU["LocPinX"].Formula);
            Assert.AreEqual("Height*0.5", test_shape_4.CellsU["LocPinY"].Formula);
            Assert.AreEqual("1 in.", test_shape_4.CellsU["Width"].Formula);
            Assert.AreEqual("0.75 in.", test_shape_4.CellsU["Height"].Formula);
            Assert.AreEqual("0 deg.", test_shape_4.CellsU["Angle"].Formula);
            Assert.AreEqual("Alone", test_shape_4.Text);

            // Sheet.20 
            var test_shape_5 = page.Shapes["Sheet.20"];
            //Assert.AreEqual("11.523988095238 in.", test_shape_5.CellsU["PinX"].Formula);
            //Assert.AreEqual("9.6104761904762 in.", test_shape_5.CellsU["PinY"].Formula);
            Assert.AreEqual("Width*0.5", test_shape_5.CellsU["LocPinX"].Formula);
            Assert.AreEqual("Height*0.5", test_shape_5.CellsU["LocPinY"].Formula);
            Assert.AreEqual("0 in.", test_shape_5.CellsU["Width"].Formula);
            Assert.AreEqual("2.6428571428571 in.", test_shape_5.CellsU["Height"].Formula);
            Assert.AreEqual("0 deg.", test_shape_5.CellsU["Angle"].Formula);
            Assert.AreEqual("", test_shape_5.Text);

            // Sheet.21 
            var test_shape_6 = page.Shapes["Sheet.21"];
            //Assert.AreEqual("9.0918452380952 in.", test_shape_6.CellsU["PinX"].Formula);
            //Assert.AreEqual("6.0287491846054 in.", test_shape_6.CellsU["PinY"].Formula);
            Assert.AreEqual("Width*0.5", test_shape_6.CellsU["LocPinX"].Formula);
            Assert.AreEqual("Height*0.5", test_shape_6.CellsU["LocPinY"].Formula);
            Assert.AreEqual("4.8642857142857 in.", test_shape_6.CellsU["Width"].Formula);
            Assert.AreEqual("3.0205968688845 in.", test_shape_6.CellsU["Height"].Formula);
            Assert.AreEqual("0 deg.", test_shape_6.CellsU["Angle"].Formula);
            Assert.AreEqual("", test_shape_6.Text);

            // Sheet.22 
            var test_shape_7 = page.Shapes["Sheet.22"];
            //Assert.AreEqual("15.309702380952 in.", test_shape_7.CellsU["PinX"].Formula);
            //Assert.AreEqual("9.7832981601732 in.", test_shape_7.CellsU["PinY"].Formula);
            Assert.AreEqual("Width*0.5", test_shape_7.CellsU["LocPinX"].Formula);
            Assert.AreEqual("Height*0.5", test_shape_7.CellsU["LocPinY"].Formula);
            Assert.AreEqual("6.5714285714286 in.", test_shape_7.CellsU["Width"].Formula);
            Assert.AreEqual("3.5585768398268 in.", test_shape_7.CellsU["Height"].Formula);
            Assert.AreEqual("0 deg.", test_shape_7.CellsU["Angle"].Formula);
            Assert.AreEqual("", test_shape_7.Text);

            // Sheet.23 
            var test_shape_8 = page.Shapes["Sheet.23"];
            //Assert.AreEqual("8.1668452380952 in.", test_shape_8.CellsU["PinX"].Formula);
            //Assert.AreEqual("8.0801435094586 in.", test_shape_8.CellsU["PinY"].Formula);
            Assert.AreEqual("Width*0.5", test_shape_8.CellsU["LocPinX"].Formula);
            Assert.AreEqual("Height*0.5", test_shape_8.CellsU["LocPinY"].Formula);
            Assert.AreEqual("3.7142857142857 in.", test_shape_8.CellsU["Width"].Formula);
            Assert.AreEqual("6.3679060665362 in.", test_shape_8.CellsU["Height"].Formula);
            Assert.AreEqual("0 deg.", test_shape_8.CellsU["Angle"].Formula);
            Assert.AreEqual("", test_shape_8.Text);

            // Sheet.24 
            var test_shape_9 = page.Shapes["Sheet.24"];
            //Assert.AreEqual("3.9525595238095 in.", test_shape_9.CellsU["PinX"].Formula);
            //Assert.AreEqual("2.4685975212003 in.", test_shape_9.CellsU["PinY"].Formula);
            Assert.AreEqual("Width*0.5", test_shape_9.CellsU["LocPinX"].Formula);
            Assert.AreEqual("Height*0.5", test_shape_9.CellsU["LocPinY"].Formula);
            Assert.AreEqual("4.7142857142857 in.", test_shape_9.CellsU["Width"].Formula);
            Assert.AreEqual("2.8551859099804 in.", test_shape_9.CellsU["Height"].Formula);
            Assert.AreEqual("0 deg.", test_shape_9.CellsU["Angle"].Formula);
            Assert.AreEqual("", test_shape_9.Text);

            // Sheet.25 
            var test_shape_10 = page.Shapes["Sheet.25"];
            //Assert.AreEqual("5.5597023809524 in.", test_shape_10.CellsU["PinX"].Formula);
            //Assert.AreEqual("6.4256669928245 in.", test_shape_10.CellsU["PinY"].Formula);
            Assert.AreEqual("Width*0.5", test_shape_10.CellsU["LocPinX"].Formula);
            Assert.AreEqual("Height*0.5", test_shape_10.CellsU["LocPinY"].Formula);
            Assert.AreEqual("8.9285714285714 in.", test_shape_10.CellsU["Width"].Formula);
            Assert.AreEqual("10.344667318982 in.", test_shape_10.CellsU["Height"].Formula);
            Assert.AreEqual("0 deg.", test_shape_10.CellsU["Angle"].Formula);
            Assert.AreEqual("", test_shape_10.Text);

            // Sheet.26 
            var test_shape_11 = page.Shapes["Sheet.26"];
            //Assert.AreEqual("12.023988095238 in.", test_shape_11.CellsU["PinX"].Formula);
            //Assert.AreEqual("6.2176190476191 in.", test_shape_11.CellsU["PinY"].Formula);
            Assert.AreEqual("Width*0.5", test_shape_11.CellsU["LocPinX"].Formula);
            Assert.AreEqual("Height*0.5", test_shape_11.CellsU["LocPinY"].Formula);
            Assert.AreEqual("1 in.", test_shape_11.CellsU["Width"].Formula);
            Assert.AreEqual("0.5 in.", test_shape_11.CellsU["Height"].Formula);
            Assert.AreEqual("0 deg.", test_shape_11.CellsU["Angle"].Formula);
            Assert.AreEqual("YES", test_shape_11.Text);

            // Sheet.27 
            var test_shape_12 = page.Shapes["Sheet.27"];
            //Assert.AreEqual("21.45255952381 in.", test_shape_12.CellsU["PinX"].Formula);
            //Assert.AreEqual("9.6104761904762 in.", test_shape_12.CellsU["PinY"].Formula);
            Assert.AreEqual("Width*0.5", test_shape_12.CellsU["LocPinX"].Formula);
            Assert.AreEqual("Height*0.5", test_shape_12.CellsU["LocPinY"].Formula);
            Assert.AreEqual("1 in.", test_shape_12.CellsU["Width"].Formula);
            Assert.AreEqual("0.5 in.", test_shape_12.CellsU["Height"].Formula);
            Assert.AreEqual("0 deg.", test_shape_12.CellsU["Angle"].Formula);
            Assert.AreEqual("NO", test_shape_12.Text); 



            doc.Close(true);

        }

        [TestMethod]
        public void RenderDirectedGraphWithDynamicConnectors()
        {
            var directed_graph_drawing = new VA.Layout.Models.DirectedGraph.Drawing();

            var n0 = directed_graph_drawing.AddShape("n0", "Untitled Node", "basflo_u.vss",
                                   "Decision");
            n0.Size = new VA.Drawing.Size(3, 2);
            var n1 = directed_graph_drawing.AddShape("n1", "", "basflo_u.vss",
                                   "Decision");

            n1.Cells = new VA.DOM.ShapeCells();
            n1.Cells.FillForegnd = "rgb(255,0,0)";
            n1.Cells.FillBkgnd = "rgb(255,255,0)";
            n1.Cells.FillPattern = 40;
            var n2 = directed_graph_drawing.AddShape("n2", "MailServer", "server_u.vss",
                                   "Email Server");
            var n3 = directed_graph_drawing.AddShape("n3", null, "basflo_u.vss",
                                   "Data");
            var n4 = directed_graph_drawing.AddShape("n4", "Alone", "basflo_u.vss",
                                   "Data");

            var c0 = directed_graph_drawing.Connect("c0", n0, n1, null, VA.Connections.ConnectorType.Curved);
            var c1 = directed_graph_drawing.Connect("c1", n1, n2, "YES", VA.Connections.ConnectorType.RightAngle);
            var c3 = directed_graph_drawing.Connect("c2", n1, n0, "NO", VA.Connections.ConnectorType.Curved);
            var c4 = directed_graph_drawing.Connect("c3", n0, n2, null, VA.Connections.ConnectorType.Straight);
            var c5 = directed_graph_drawing.Connect("c4", n2, n3, null, VA.Connections.ConnectorType.Curved);
            var c6 = directed_graph_drawing.Connect("c5", n3, n0, null, VA.Connections.ConnectorType.Curved);

            var options = new VA.Layout.MSAGL.LayoutOptions();
            options.UseDynamicConnectors = true;

            var visapp = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page1 = visapp.ActivePage;
            VA.Layout.MSAGL.MSAGLRenderer.Render(page1, directed_graph_drawing, options);

            Assert.IsNotNull(n0.VisioShape);
            Assert.IsNotNull(n1.VisioShape);
            Assert.IsNotNull(n2.VisioShape);
            Assert.IsNotNull(n3.VisioShape);
            Assert.IsNotNull(n4.VisioShape);

            Assert.AreEqual("Untitled Node", n0.VisioShape.Text);
            Assert.AreEqual("", n1.VisioShape.Text);
            Assert.AreEqual("MailServer", n2.VisioShape.Text);
            Assert.AreEqual("", n3.VisioShape.Text);
            Assert.AreEqual("Alone", n4.VisioShape.Text);

            Assert.AreEqual(new VA.Drawing.Size(3, 2), VisioAutomationTest.GetSize(n0.VisioShape));
            Assert.AreEqual(options.DefaultShapeSize, VisioAutomationTest.GetSize(n1.VisioShape));
            Assert.AreEqual("40", n1.VisioShape.CellsU["FillPattern"].FormulaU);

            Assert.IsNotNull(c0.VisioShape);
            Assert.IsNotNull(c1.VisioShape);
            Assert.IsNotNull(c3.VisioShape);
            Assert.IsNotNull(c4.VisioShape);
            Assert.IsNotNull(c5.VisioShape);
            Assert.IsNotNull(c6.VisioShape);

            Assert.AreEqual("", c0.VisioShape.Text);
            Assert.AreEqual("YES", c1.VisioShape.Text);
            Assert.AreEqual("NO", c3.VisioShape.Text);
            Assert.AreEqual("", c4.VisioShape.Text);
            Assert.AreEqual("", c5.VisioShape.Text);
            Assert.AreEqual("", c6.VisioShape.Text);

            var pagesize = page1.GetSize();
            //TestUtil.AreEqual(13.62, 13.57, pagesize, 0.05);

            doc.Close(true);
        }

        [TestMethod]
        public void RenderDirectedGraphWithCustomProps()
        {
            var directed_graph_drawing = new VA.Layout.Models.DirectedGraph.Drawing();

            var n0 = directed_graph_drawing.AddShape("n0", "Untitled Node", "basflo_u.vss",
                                   "Decision");
            n0.Size = new VA.Drawing.Size(3, 2);
            n0.CustomProperties = new Dictionary<string, VA.CustomProperties.CustomPropertyCells>();
            n0.CustomProperties["p1"] = new VA.CustomProperties.CustomPropertyCells("v1");
            n0.CustomProperties["p2"] = new VA.CustomProperties.CustomPropertyCells("v2");
            n0.CustomProperties["p3"] = new VA.CustomProperties.CustomPropertyCells("v3");

            var options = new VA.Layout.MSAGL.LayoutOptions();
            options.UseDynamicConnectors = true;

            var visapp = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page1 = visapp.ActivePage;
            VA.Layout.MSAGL.MSAGLRenderer.Render(page1, directed_graph_drawing, options);
            
            Assert.IsNotNull(n0.VisioShape);
            var props_dic = VA.CustomProperties.CustomPropertyHelper.GetCustomProperties(n0.VisioShape);
            Assert.IsTrue(props_dic.Count>=3);
            Assert.AreEqual("\"v1\"",props_dic["p1"].Value.Formula);
            Assert.AreEqual("\"v2\"", props_dic["p2"].Value.Formula);
            Assert.AreEqual("\"v3\"", props_dic["p3"].Value.Formula);

            page1.Application.ActiveWindow.ViewFit = (short) IVisio.VisWindowFit.visFitPage;

            doc.Close(true);
        }
    }
}