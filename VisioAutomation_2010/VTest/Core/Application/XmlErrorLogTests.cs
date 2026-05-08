using System.IO;
using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Application.Logging;

namespace VTest.Core.Application
{

    [MUT.TestClass]
    public class XmlErrorLogTests : Framework.VTest
    {
        [MUT.TestMethod]
        public void XmlErrorLog_Visio2013VSDXLog_Reads51Sessions()
        {
            string input_filename = this._get_test_results_out_path(@"datafiles\VSDX_Log_Visio_2013.txt");

            MUT.Assert.IsTrue(File.Exists(input_filename));
            var log = new XmlErrorLog(input_filename);
            MUT.Assert.AreEqual(51, log.LogSessions.Count);


        }

        [MUT.TestMethod]
        public void XmlErrorLog_Visio2010Sample_ParsesSessionsAndDataTypeWarnings()
        {
            string input_filename = this._get_test_results_out_path(@"datafiles\XMLErrorLog_Visio_2010_1.txt");

            MUT.Assert.IsTrue(File.Exists(input_filename));
            var log = new XmlErrorLog(input_filename);
            MUT.Assert.AreEqual(2,log.LogSessions.Count);

            var first_session = log.LogSessions[0];
            var second_session = log.LogSessions[1];

            MUT.Assert.IsTrue(first_session.Source.EndsWith("vdx_with_warnings_1.vdx"));
            MUT.Assert.IsTrue(second_session.Source.EndsWith("VDX_Tests.VDX_MultiPageDocument2015-10-1--20-09-10.vdx"));

            MUT.Assert.AreEqual(4, first_session.LogRecords.Count);
            MUT.Assert.AreEqual(2, second_session.LogRecords.Count);

            MUT.Assert.IsTrue(first_session.LogRecords[0].Type == "Warning" && first_session.LogRecords[0].SubType=="DataType");
            MUT.Assert.IsTrue(first_session.LogRecords[1].Type == "Warning" && first_session.LogRecords[1].SubType == "DataType");
            MUT.Assert.IsTrue(first_session.LogRecords[2].Type == "Warning" && first_session.LogRecords[2].SubType == "DataType");
            MUT.Assert.IsTrue(first_session.LogRecords[3].Type == "Warning" && first_session.LogRecords[3].SubType == "DataType");

            MUT.Assert.IsTrue(second_session.LogRecords[0].Type == "Warning" && second_session.LogRecords[0].SubType == "DataType");
            MUT.Assert.IsTrue(second_session.LogRecords[1].Type == "Warning" && second_session.LogRecords[1].SubType == "DataType");
        }

        [MUT.TestMethod]
        public void XmlErrorLog_Visio2013Sample_ParsesFourSessionsWithWarningsInLast()
        {
            string input_filename = this._get_test_results_out_path(@"datafiles\XMLErrorLog_Visio_2013_1.txt");

            MUT.Assert.IsTrue(File.Exists(input_filename));
            var log = new XmlErrorLog(input_filename);
            MUT.Assert.AreEqual(4, log.LogSessions.Count);

            var first_session = log.LogSessions[0];
            var second_session = log.LogSessions[1];
            var third_session = log.LogSessions[2];
            var fourth_session = log.LogSessions[3];

            MUT.Assert.AreEqual(0, first_session.LogRecords.Count);
            MUT.Assert.AreEqual(0, second_session.LogRecords.Count);
            MUT.Assert.AreEqual(0, third_session.LogRecords.Count);
            MUT.Assert.AreEqual(2, fourth_session.LogRecords.Count);

            MUT.Assert.IsTrue(first_session.Source.EndsWith("template_router.vdx"));
            MUT.Assert.IsTrue(second_session.Source.EndsWith("COMPS_U.VSSX"));
            MUT.Assert.IsTrue(third_session.Source.EndsWith("PERIPH_U.VSSX"));
            MUT.Assert.IsTrue(fourth_session.Source.EndsWith("vdx_with_warnings_1.vdx"));

            MUT.Assert.IsTrue(first_session.LogRecords.Count==0);
        }
    }
}
