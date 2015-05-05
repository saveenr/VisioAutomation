using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TestVisioAutomation.ImportExport
{

    [TestClass]
    public class XmlErrorLog_Tests2 : VisioAutomationTest
    {
        [TestMethod]
        [DeploymentItem(@"datafiles\VSDX_Log_Visio_2013.txt", "datafiles")]
        public void VSD_Load_Visio2013()
        {
            string input_filename = this.GetTestResultsOutPath(@"datafiles\VSDX_Log_Visio_2013.txt");

            Assert.IsTrue(File.Exists(input_filename));
            var log = new VisioAutomation.Application.Logging.XmlErrorLog(input_filename);
            Assert.AreEqual(51, log.FileSessions.Count);


        }
    }

    [TestClass]
    public class XmlErrorLog_Tests : VisioAutomationTest
    {
        [TestMethod]
        [DeploymentItem(@"datafiles\XMLErrorLog_Visio_2010_1.txt", "datafiles")]
        public void XmlErrorLog_Load_Visio2010_1()
        {
            string input_filename = this.GetTestResultsOutPath(@"datafiles\XMLErrorLog_Visio_2010_1.txt");

            Assert.IsTrue(File.Exists(input_filename));
            var log = new VisioAutomation.Application.Logging.XmlErrorLog(input_filename);
            Assert.AreEqual(2,log.FileSessions.Count);

            var first_session = log.FileSessions[0];
            var second_session = log.FileSessions[1];

            Assert.IsTrue(first_session.Source.EndsWith("vdx_with_warnings_1.vdx"));
            Assert.IsTrue(second_session.Source.EndsWith("VDX_Tests.VDX_MultiPageDocument2015-10-1--20-09-10.vdx"));

            Assert.AreEqual(4, first_session.Records.Count);
            Assert.AreEqual(2, second_session.Records.Count);

            Assert.IsTrue(first_session.Records[0].Type == "Warning" && first_session.Records[0].SubType=="DataType");
            Assert.IsTrue(first_session.Records[1].Type == "Warning" && first_session.Records[1].SubType == "DataType");
            Assert.IsTrue(first_session.Records[2].Type == "Warning" && first_session.Records[2].SubType == "DataType");
            Assert.IsTrue(first_session.Records[3].Type == "Warning" && first_session.Records[3].SubType == "DataType");

            Assert.IsTrue(second_session.Records[0].Type == "Warning" && second_session.Records[0].SubType == "DataType");
            Assert.IsTrue(second_session.Records[1].Type == "Warning" && second_session.Records[1].SubType == "DataType");
        }

        [TestMethod]
        [DeploymentItem(@"datafiles\XMLErrorLog_Visio_2013_1.txt", "datafiles")]
        public void XmlErrorLog_Load_Visio2013_1()
        {
            string input_filename = this.GetTestResultsOutPath(@"datafiles\XMLErrorLog_Visio_2013_1.txt");

            Assert.IsTrue(File.Exists(input_filename));
            var log = new VisioAutomation.Application.Logging.XmlErrorLog(input_filename);
            Assert.AreEqual(4, log.FileSessions.Count);

            var first_session = log.FileSessions[0];
            var second_session = log.FileSessions[1];
            var third_session = log.FileSessions[2];
            var fourth_session = log.FileSessions[3];

            Assert.AreEqual(0, first_session.Records.Count);
            Assert.AreEqual(0, second_session.Records.Count);
            Assert.AreEqual(0, third_session.Records.Count);
            Assert.AreEqual(2, fourth_session.Records.Count);

            Assert.IsTrue(first_session.Source.EndsWith("template_router.vdx"));
            Assert.IsTrue(second_session.Source.EndsWith("COMPS_U.VSSX"));
            Assert.IsTrue(third_session.Source.EndsWith("PERIPH_U.VSSX"));
            Assert.IsTrue(fourth_session.Source.EndsWith("vdx_with_warnings_1.vdx"));

            Assert.IsTrue(first_session.Records.Count==0);
        }
    }
}
