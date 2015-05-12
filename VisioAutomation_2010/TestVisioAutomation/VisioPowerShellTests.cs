
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Data;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
 
namespace TestVisioAutomation
{
    [TestClass]
    public class VisioPowerShellTests
    {

        //http://nivot.org/blog/post/2010/05/03/PowerShell20DeveloperEssentials1InitializingARunspaceWithAModule
	 
	    // This is needed so the VisioPS.dll is copied to the "Test Results\Out" directory... the directory where the tests are "running" from
        // https://connect.microsoft.com/VisualStudio/feedback/details/771138/vs2012-referenced-assemblies-in-unit-test-are-not-copied-to-the-unit-test-out-f
        VisioPowerShell.Commands.New_VisioApplication visioApp = null;
	 
        private static PowerShell powerShell;
        private static InitialSessionState sessionState;
        private static Runspace runSpace;
        private static RunspaceInvoke invoker;
 
        [ClassInitialize]
        public static void PSTestFixtureSetup(TestContext context)
        {
            var test = new VisioPowerShell.Commands.New_VisioApplication().Invoke();

            sessionState = InitialSessionState.CreateDefault();

            
            // Get path of where everything is executing so we can find the VisioPS.dll assembly
            var executing_assembly = System.Reflection.Assembly.GetExecutingAssembly();
            var path = System.IO.Path.GetDirectoryName(executing_assembly.GetName().CodeBase);
            var uri = new Uri(path);
            var visioPS = uri.LocalPath + "\\VisioPS.dll";
            // Import the latest VisioPS module into the PowerShell session
            sessionState.ImportPSModule(new []{visioPS});
            runSpace = RunspaceFactory.CreateRunspace(sessionState);
            runSpace.Open();
            powerShell = PowerShell.Create();
            powerShell.Runspace = runSpace;
            invoker = new RunspaceInvoke(runSpace);
        }
 
        [TestCleanup]
        public void PSTestFixtureTeardown()
        {
           
        }
 
        [ClassCleanup]
        public static void PSTestVisioPowerShellClassCleanup()
        {
            // Make sure we cleanup everything
            powerShell.Dispose();
            invoker.Dispose();
            runSpace.Close();
 
            invoker = null;
            runSpace = null;
            sessionState = null;
            powerShell = null;
        }
 
        [TestMethod]
        public void PSTestNewVisioDocument()
        {
           // Send the command to the PowerShell session
            var visDoc = invoker.Invoke("New-VisioDocument");
 
            // Verify results
            Assert.IsNotNull(visDoc);
            Assert.AreEqual("Microsoft.Office.Interop.Visio.DocumentClass", visDoc[0].ToString());
            Assert.IsNotNull(visDoc[0].Properties["Name"].Value);
            Assert.IsFalse(String.IsNullOrEmpty(visDoc[0].Properties["Name"].Value.ToString()));
 
            // Close Visio Application that was created when "New-VisioDocument" was invoked
            invoker.Invoke("Close-VisioApplication");
        }
 
        [TestMethod]
        public void PSTestGetVisioPageCell()
        {
            var visDoc = invoker.Invoke("New-VisioDocument");
            var visGetPageCell = invoker.Invoke("Get-VisioPageCell -Cells PageWidth,PageHeight -Page (Get-VisioPage -ActivePage) -GetResults -ResultType Double");
            DataRow results = ((DataRowCollection)visGetPageCell[0].Properties["Rows"].Value)[0];
           powerShell = PowerShell.Create(); ;
            Assert.IsNotNull(visGetPageCell);
            Assert.AreEqual(8.5, results["PageWidth"]);
            Assert.AreEqual(11.0, results["PageHeight"]);
            
            //Now lets add another page and get it's width and height
            var page2 = invoker.Invoke("New-VisioPage");
            var visGetPageCell2 = invoker.Invoke("Get-VisioPageCell -Cells PageWidth,PageHeight -Page (Get-VisioPage -ActivePage) -GetResults -ResultType Double");
            DataRow results2 = ((DataRowCollection)visGetPageCell2[0].Properties["Rows"].Value)[0];
 
            Assert.IsNotNull(visGetPageCell2);
 	        Assert.AreEqual(8.5, results2["PageWidth"]);
	        Assert.AreEqual(11.0, results2["PageHeight"]);

            // Close Visio Application that was created when "New-VisioDocument" was invoked
            invoker.Invoke("Close-VisioApplication -Force");
        }

      [TestMethod]
      public void NewVisioContainer()
      {
          var visDoc = invoker.Invoke("New-VisioDocument");

          var rect = invoker.Invoke("(Get-VisioMaster \"Rectangle\" (Open-VisioDocument \"BASIC_U.VSS\"))");

          // Another way to send a command...
          Pipeline pipeline = runSpace.CreatePipeline();
          
          Command myCmd = new Command(@"New-VisioShape");
          CommandParameter myCmd1 = new CommandParameter("Masters", rect);
          myCmd.Parameters.Add(myCmd1);
          
          double[] points = { 1, 1 };
          CommandParameter myCmd2 = new CommandParameter("Points", points);
          myCmd.Parameters.Add(myCmd2);

          pipeline.Commands.Add(myCmd);
          pipeline.Invoke();

          // Everything above (to the new "pipeline" variable) can be done with this one line...
          //var shape = invoker.Invoke("New-VisioShape -Masters (Get-VisioMaster \"Rectangle\" (Open-VisioDocument \"BASIC_U.VSS\")) -Points 1,1");
          
          // Drop a container on the page... the rectangle we created above should be selected by default. 
          // Since it is selected it will be added as a member to the container.
          var container = invoker.Invoke("New-VisioContainer -Masters (Get-VisioMaster \"Container 1\" (Open-VisioDocument \"SDCONT_U.VSS\"))");

          Assert.IsNotNull(container);
          
          // Cleanup
          invoker.Invoke("Close-VisioApplication -Force");
      }
    }
}
	 