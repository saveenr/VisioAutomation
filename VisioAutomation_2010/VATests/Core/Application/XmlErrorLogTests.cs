using System.IO;
using VisioAutomation.Application.Logging;

namespace VisioAutomation_Tests.Core.Application;

[TestClass]
public class XmlErrorLogTests : VisioAutomationTest
{
    [TestMethod]
    [DeploymentItem(@"datafiles\VSDX_Log_Visio_2013.txt", "datafiles")]
    public void VSD_Load_Visio2013()
    {
        string input_filename = this._get_test_results_out_path(@"datafiles\VSDX_Log_Visio_2013.txt");

        Assert.IsTrue(File.Exists(input_filename));
        var log = new XmlErrorLog(input_filename);
        Assert.AreEqual(51, log.LogSessions.Count);


    }

    [TestMethod]
    [DeploymentItem(@"datafiles\XMLErrorLog_Visio_2010_1.txt", "datafiles")]
    public void XmlErrorLog_Load_Visio2010_1()
    {
        string input_filename = this._get_test_results_out_path(@"datafiles\XMLErrorLog_Visio_2010_1.txt");

        Assert.IsTrue(File.Exists(input_filename));
        var log = new XmlErrorLog(input_filename);
        Assert.AreEqual(2,log.LogSessions.Count);

        var first_session = log.LogSessions[0];
        var second_session = log.LogSessions[1];

        Assert.IsTrue(first_session.Source.EndsWith("vdx_with_warnings_1.vdx"));
        Assert.IsTrue(second_session.Source.EndsWith("VDX_Tests.VDX_MultiPageDocument2015-10-1--20-09-10.vdx"));

        Assert.AreEqual(4, first_session.LogRecords.Count);
        Assert.AreEqual(2, second_session.LogRecords.Count);

        Assert.IsTrue(first_session.LogRecords[0].Type == "Warning" && first_session.LogRecords[0].SubType=="DataType");
        Assert.IsTrue(first_session.LogRecords[1].Type == "Warning" && first_session.LogRecords[1].SubType == "DataType");
        Assert.IsTrue(first_session.LogRecords[2].Type == "Warning" && first_session.LogRecords[2].SubType == "DataType");
        Assert.IsTrue(first_session.LogRecords[3].Type == "Warning" && first_session.LogRecords[3].SubType == "DataType");

        Assert.IsTrue(second_session.LogRecords[0].Type == "Warning" && second_session.LogRecords[0].SubType == "DataType");
        Assert.IsTrue(second_session.LogRecords[1].Type == "Warning" && second_session.LogRecords[1].SubType == "DataType");
    }

    [TestMethod]
    [DeploymentItem(@"datafiles\XMLErrorLog_Visio_2013_1.txt", "datafiles")]
    public void XmlErrorLog_Load_Visio2013_1()
    {
        string input_filename = this._get_test_results_out_path(@"datafiles\XMLErrorLog_Visio_2013_1.txt");

        Assert.IsTrue(File.Exists(input_filename));
        var log = new XmlErrorLog(input_filename);
        Assert.AreEqual(4, log.LogSessions.Count);

        var first_session = log.LogSessions[0];
        var second_session = log.LogSessions[1];
        var third_session = log.LogSessions[2];
        var fourth_session = log.LogSessions[3];

        Assert.AreEqual(0, first_session.LogRecords.Count);
        Assert.AreEqual(0, second_session.LogRecords.Count);
        Assert.AreEqual(0, third_session.LogRecords.Count);
        Assert.AreEqual(2, fourth_session.LogRecords.Count);

        Assert.IsTrue(first_session.Source.EndsWith("template_router.vdx"));
        Assert.IsTrue(second_session.Source.EndsWith("COMPS_U.VSSX"));
        Assert.IsTrue(third_session.Source.EndsWith("PERIPH_U.VSSX"));
        Assert.IsTrue(fourth_session.Source.EndsWith("vdx_with_warnings_1.vdx"));

        Assert.IsTrue(first_session.LogRecords.Count==0);
    }
}