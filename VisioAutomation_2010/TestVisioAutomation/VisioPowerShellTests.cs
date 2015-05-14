using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using SMA=System.Management.Automation;
using IVisio=Microsoft.Office.Interop.Visio;
using System.Collections.Generic;

namespace TestVisioAutomation
{
    [TestClass]
    public class VisioPowerShellTests
    {
        private static PowerShellContext ps_context = new PowerShellContext();

        [ClassInitialize]
        public static void PSTestFixtureSetup(TestContext context)
        {
            var visio_app = new VisioPowerShell.Commands.New_VisioApplication().Invoke();
        }
 
        [TestCleanup]
        public void PSTestFixtureTeardown()
        {
           
        }
 
        [ClassCleanup]
        public static void PSTestVisioPowerShellClassCleanup()
        {
            VisioPowerShellTests.ps_context.CleanUp();
        }
 
        [TestMethod]
        public void PSTestNewVisioDocument()
        {
            var doc = VisioPowerShellTests.New_Visio_Document();
 
            // Verify results
            Assert.IsNotNull(doc);
            var ps_object = doc[0];
            Assert.AreEqual("Microsoft.Office.Interop.Visio.DocumentClass", ps_object.ToString());
            Assert.IsNotNull(ps_object.Properties["Name"].Value);
            Assert.IsFalse(String.IsNullOrEmpty(ps_object.Properties["Name"].Value.ToString()));
 
            // Close Visio Application that was created when "New-VisioDocument" was invoked
            VisioPowerShellTests.Close_Visio_Application();
        }

        private static void Close_Visio_Application()
        {
            VisioPowerShellTests.ps_context.Invoker.Invoke("Close-VisioApplication -Force");
        }

        [TestMethod]
        public void PSTestGetVisioPageCell()
        {
            var doc = VisioPowerShellTests.New_Visio_Document();
            var cells1 = VisioPowerShellTests.ps_context.Invoker.Invoke("Get-VisioPageCell -Cells PageWidth,PageHeight -Page (Get-VisioPage -ActivePage) -GetResults -ResultType Double");
            var data_row_collection1 = (System.Data.DataRowCollection)cells1[0].Properties["Rows"].Value;
            var results = data_row_collection1[0];
            Assert.IsNotNull(cells1);
            Assert.AreEqual(8.5, results["PageWidth"]);
            Assert.AreEqual(11.0, results["PageHeight"]);
            
            //Now lets add another page and get it's width and height
            var page2 = VisioPowerShellTests.ps_context.Invoker.Invoke("New-VisioPage");
            var cells2 = VisioPowerShellTests.ps_context.Invoker.Invoke("Get-VisioPageCell -Cells PageWidth,PageHeight -Page (Get-VisioPage -ActivePage) -GetResults -ResultType Double");
            var data_row_collection2 = (System.Data.DataRowCollection)cells2[0].Properties["Rows"].Value;
            var results2 = data_row_collection2[0];
 
            Assert.IsNotNull(cells2);
 	        Assert.AreEqual(8.5, results2["PageWidth"]);
	        Assert.AreEqual(11.0, results2["PageHeight"]);

            VisioPowerShellTests.Close_Visio_Application();
        }


      [TestMethod]
      public void PSNewVisioContainer()
      {
          var doc = VisioPowerShellTests.New_Visio_Document();
          var app = VisioPowerShellTests.Get_Visio_Application();

          var ver = VisioAutomation.Application.ApplicationHelper.GetVersion(app);

          var cont_doc = ver.Major >= 15 ? "SDCONT_U.VSSX" : "SDCONT_U.VSS";
          var cont_master_name = ver.Major >= 15 ? "Plain" : "Container 1";
          var rectangle = "Rectangle";
          var basic_u_vss = "BASIC_U.VSS";

          var rect = VisioPowerShellTests.Get_Visio_Master(rectangle, basic_u_vss);

          VisioPowerShellTests.New_VisioShape(rect, new[] {1.0, 1.0});

          // Drop a container on the page... the rectangle we created above should be selected by default. 
          // Since it is selected it will be added as a member to the container.

          var container = VisioPowerShellTests.New_Visio_Container(cont_master_name, cont_doc);

          Assert.IsNotNull(container);

          VisioPowerShellTests.Close_Visio_Application();
      }

        private static IVisio.ShapeClass New_Visio_Container(string cont_master_name, string cont_doc)
        {
            var cmd = string.Format("New-VisioContainer -Master (Get-VisioMaster \"{0}\" (Open-VisioDocument \"{1}\"))", cont_master_name, cont_doc);
            var results = VisioPowerShellTests.ps_context.Invoker.Invoke(cmd);
            var shape = (IVisio.ShapeClass)results[0].BaseObject;
            return shape;
        }

        private static List<IVisio.Shape> New_VisioShape(IVisio.MasterClass master, double[] points)
        {
            var pipeline = VisioPowerShellTests.ps_context.RunSpace.CreatePipeline();
            var cmd = new SMA.Runspaces.Command(@"New-VisioShape");
            cmd.AddParameter("Master", master);
            cmd.AddParameter("Points", points);
            pipeline.Commands.Add(cmd);
            var results = pipeline.Invoke();
            var shapes = (List<IVisio.Shape>)results[0].BaseObject;
            return shapes;
        }

        private static IVisio.MasterClass Get_Visio_Master(string rectangle, string basic_u_vss)
        {
            var cmd = string.Format("(Get-VisioMaster \"{0}\" (Open-VisioDocument \"{1}\"))", rectangle, basic_u_vss);
            var results = VisioPowerShellTests.ps_context.Invoker.Invoke(cmd);
            var master = (IVisio.MasterClass)results[0].BaseObject;
            return master;
        }

        private static System.Collections.ObjectModel.Collection<System.Management.Automation.PSObject> New_Visio_Document()
        {
            var doc = VisioPowerShellTests.ps_context.Invoker.Invoke("New-VisioDocument");
            return doc;
        }

        private static Microsoft.Office.Interop.Visio.ApplicationClass Get_Visio_Application()
        {
            var app_0 = VisioPowerShellTests.ps_context.Invoker.Invoke("Get-VisioApplication");
            var app = (IVisio.ApplicationClass) app_0[0].BaseObject;
            return app;
        }
    }

    public static class SMA_Extensions
    {
        public static void AddParameter(this SMA.Runspaces.Command cmd, string name, object value)
        {
            var parameter= new SMA.Runspaces.CommandParameter(name, value);
            cmd.Parameters.Add(parameter);            
        }    
    }
}
	 