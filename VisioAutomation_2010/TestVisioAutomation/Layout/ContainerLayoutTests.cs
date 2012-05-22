using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using OCMODEL = VisioAutomation.Layout.Models.ContainerLayout;

namespace TestVisioAutomation
{
    [TestClass]
    public class CointainerLayoutTests : VisioAutomationTest
    {
        [TestMethod]
        public void ContainerMustCallPerformLayout()
        {
            // Purpose: Verify that if PerformLayout is NOT called before Render() 
            // is called then an exception will be thrown

            bool caught = false;
            var layout = new OCMODEL.ContainerLayout();
            var doc = this.GetNewDoc();
            IVisio.Page page = null;
            try
            {
                var c1 = layout.AddContainer("A");
                var i1 = c1.Add("A1");
                // layout.PerformLayout(); 
                page = layout.Render(doc);
                page.Delete(0);
            }
            catch (VA.AutomationException exc)
            {
                caught = true;
            }

            doc.Close(true);

            if (caught == false)
            {
                Assert.Fail("Did not catch expected exception");
            }
        }

        [TestMethod]
        public void DrawContainer1()
        {

            // Purpose: Simple test to make sure that both Containers and Non-Container
            // rendering are supported. The diagram is a single container having a single
            // container item

            var doc = this.GetNewDoc();

            var layout1 = new OCMODEL.ContainerLayout();
            var l1_c1 = layout1.AddContainer("L1/C1");
            var l1_c1_i1 = l1_c1.Add("L1/C1/I1");
            
            layout1.PerformLayout();

            layout1.LayoutOptions.Style = VA.Layout.Models.ContainerLayout.RenderStyle.UseVisioContainers;
            var page1 = layout1.Render(doc);
            page1.Delete(0);

            layout1.LayoutOptions.Style = VA.Layout.Models.ContainerLayout.RenderStyle.UseShapes;
            page1 = layout1.Render(doc);

            page1.Delete(0);

            doc.Close(true);
        }
    }
}