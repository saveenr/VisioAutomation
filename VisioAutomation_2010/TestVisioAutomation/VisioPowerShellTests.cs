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
            var doc = ps_context.New_Visio_Document();
            Assert.IsNotNull(doc);
            VisioPowerShellTests.Close_Visio_Application();
        }

        private static void Close_Visio_Application()
        {
            VisioPowerShellTests.ps_context.Close_Visio_Application();
        }

        [TestMethod]
        public void PSTestGetVisioPageCell()
        {
            var doc = ps_context.New_Visio_Document();
            var datatable1 = ps_context.Get_Visio_Page_Cell(new[] { "PageWidth", "PageHeight" }, true, "Double");
            var results = datatable1.Rows[0];
            Assert.IsNotNull(datatable1);
            Assert.AreEqual(8.5, results["PageWidth"]);
            Assert.AreEqual(11.0, results["PageHeight"]);
            
            //Now lets add another page and get it's width and height
            var page2 = VisioPowerShellTests.ps_context.New_Visio_Page();
            var datatable2 = ps_context.Get_Visio_Page_Cell(new[] { "PageWidth", "PageHeight" }, true, "Double");
            var results2 = datatable1.Rows[0];
 
            Assert.IsNotNull(datatable2);
 	        Assert.AreEqual(8.5, results2["PageWidth"]);
	        Assert.AreEqual(11.0, results2["PageHeight"]);

            VisioPowerShellTests.Close_Visio_Application();
        }


      [TestMethod]
      public void PSNewVisioContainer()
      {
          var doc = ps_context.New_Visio_Document();
          var app = ps_context.Get_Visio_Application();

          var ver = VisioAutomation.Application.ApplicationHelper.GetVersion(app);

          var cont_doc = ver.Major >= 15 ? "SDCONT_U.VSSX" : "SDCONT_U.VSS";
          var cont_master_name = ver.Major >= 15 ? "Plain" : "Container 1";
          var rectangle = "Rectangle";
          var basic_u_vss = "BASIC_U.VSS";

          var rect = ps_context.Get_Visio_Master(rectangle, basic_u_vss);

          ps_context.New_VisioShape(rect, new[] { 1.0, 1.0 });

          // Drop a container on the page... the rectangle we created above should be selected by default. 
          // Since it is selected it will be added as a member to the container.

          var container = ps_context.New_Visio_Container(cont_master_name, cont_doc);

          Assert.IsNotNull(container);

          VisioPowerShellTests.Close_Visio_Application();
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
	 