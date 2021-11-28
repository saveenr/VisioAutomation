using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VACONT = VisioAutomation.Models.Layouts.Container;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VTest.Models.Layouts
{
    [MUT.TestClass]
    public class CointainerLayout_Tests : Framework.VTest
    {
        [MUT.TestMethod]
        public void Container_PerformLayoutBeforeRender()
        {
            // Purpose: Verify that if PerformLayout is NOT called before Render() 
            // is called then an exception will be thrown

            bool caught = false;
            var layout = new VACONT.ContainerLayout();
            var doc = this.GetNewDoc();
            try
            {
                var c1 = layout.AddContainer("A");
                var i1 = c1.Add("A1");
                // layout.PerformLayout(); 
                IVisio.Page page = layout.Render(doc);
                page.Delete(0);
            }
            catch (System.ArgumentException)
            {
                caught = true;
            }

            doc.Close(true);

            if (caught == false)
            {
                MUT.Assert.Fail("Did not catch expected exception");
            }
        }

        [MUT.TestMethod]
        public void Container_Diagram1()
        {

            // Purpose: Simple test to make sure that both Containers and Non-Container
            // rendering are supported. The diagram is a single container having a single
            // container item

            var doc = this.GetNewDoc();

            var layout1 = new VACONT.ContainerLayout();
            var l1_c1 = layout1.AddContainer("L1/C1");
            var l1_c1_i1 = l1_c1.Add("L1/C1/I1");
            
            layout1.PerformLayout();
            var page1 = layout1.Render(doc);

            page1.Delete(0);

            doc.Close(true);
        }


        [MUT.TestMethod]
        public void Container_Diagram2()
        {
            // Make sure that empty containers can be drawn
            var doc = this.GetNewDoc();

            var layout1 = new VACONT.ContainerLayout();
            var l1_c1 = layout1.AddContainer("L1/C1");
            var l1_c1_i1 = l1_c1.Add("L1/C1/I1");
            var l1_c2 = layout1.AddContainer("L1/C2"); // this is the empty container

            layout1.PerformLayout();
            var page1 = layout1.Render(doc);

            page1.Delete(0);
            doc.Close(true);
        }
    }
}