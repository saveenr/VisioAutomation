
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
        private static PowerShell powerShell;
        private static InitialSessionState sessionState;
        private static Runspace runSpace;
        private static RunspaceInvoke invoker;
 
        [ClassInitialize]
        public static void PSTestFixtureSetup(TestContext context)
        {
            powerShell = PowerShell.Create(); ;
            sessionState = InitialSessionState.CreateDefault();
            runSpace = RunspaceFactory.CreateRunspace(sessionState);
            invoker = new RunspaceInvoke(runSpace);

            invoker.Invoke("Set-ExecutionPolicy Unrestricted");
            //string[] modules = { "Visio" };
            //sessionState.ImportPSModule(modules);
            
            // Get path of where everything is executing so we can find the VisioPS.dll assembly
            var executing_assembly = System.Reflection.Assembly.GetExecutingAssembly();
            var path = System.IO.Path.GetDirectoryName(executing_assembly.GetName().CodeBase);
            var uri = new Uri(path);
            var visioPS = uri.LocalPath + "\\VisioPS.dll";
            // Import the latest VisioPS module into the PowerShell session
            sessionState.ImportPSModulesFromPath(visioPS);
            runSpace.Open();
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
 
            Assert.IsNotNull(visGetPageCell);
            Assert.AreEqual(8.5, results["PageWidth"]);
            Assert.AreEqual(11.0, results["PageHeight"]);
            
            //Now lets add another page and get it's width and height
            var page2 = invoker.Invoke("New-VisioPage");
            var visGetPageCell2 = invoker.Invoke("Get-VisioPageCell -Cells PageWidth,PageHeight -Page (Get-VisioPage -ActivePage) -GetResults -ResultType Double");
            DataRow results2 = ((DataRowCollection)visGetPageCell2[0].Properties["Rows"].Value)[0];
 
            Assert.IsNotNull(visGetPageCell2);
 
            // Close Visio Application that was created when "New-VisioDocument" was invoked
            invoker.Invoke("Close-VisioApplication -Force");
        }
    }
}
	 