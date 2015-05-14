using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using SMA=System.Management.Automation;
using IVisio=Microsoft.Office.Interop.Visio;

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

          var cont_doc = "SDCONT_U.VSSX";
          var cont_master_name = "Plain";
          var rectangle = "Rectangle";
          var basic_u_vss = "BASIC_U.VSS";

          if (ver.Major == 14)
          {
              cont_doc = "SDCONT_U.VSS";
              cont_master_name = "Container 1";
          }

          var rect = VisioPowerShellTests.Get_Visio_Master(rectangle, basic_u_vss);

          // Another way to send a command...
          var pipeline = VisioPowerShellTests.ps_context.RunSpace.CreatePipeline();

          var cmd_1 = new SMA.Runspaces.Command(@"New-VisioShape");
          cmd_1.AddParameter("Master", rect);
          cmd_1.AddParameter("Points", new[] { 1.0, 1.0 });
          pipeline.Commands.Add(cmd_1);
          pipeline.Invoke();

          // Everything above (to the new "pipeline" variable) can be done with this one line...
          //var shape = invoker.Invoke("New-VisioShape -Master (Get-VisioMaster \"Rectangle\" (Open-VisioDocument \"BASIC_U.VSS\")) -Points 1,1");
          
          // Drop a container on the page... the rectangle we created above should be selected by default. 
          // Since it is selected it will be added as a member to the container.

          var line2 = string.Format("New-VisioContainer -Master (Get-VisioMaster \"{0}\" (Open-VisioDocument \"{1}\"))", cont_master_name, cont_doc);
          var container = VisioPowerShellTests.ps_context.Invoker.Invoke(line2);

          Assert.IsNotNull(container);

          VisioPowerShellTests.Close_Visio_Application();
      }

        private static System.Collections.ObjectModel.Collection<System.Management.Automation.PSObject> Get_Visio_Master(string rectangle, string basic_u_vss)
        {
            var line1 = string.Format("(Get-VisioMaster \"{0}\" (Open-VisioDocument \"{1}\"))", rectangle, basic_u_vss);
            var rect = VisioPowerShellTests.ps_context.Invoker.Invoke(line1);
            return rect;
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
	 