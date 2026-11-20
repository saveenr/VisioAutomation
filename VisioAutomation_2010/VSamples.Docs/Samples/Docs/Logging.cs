using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page logging.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class LoggingSamples
    {
        public static void Sample_01(IVisio.Application app)
        {
            // <sample id="logging#01" doc="logging.md" line="9">
            string path = VisioAutomation.Application.Logging.LoggingHelper
                .GetXmlErrorLogFilename(app);
            // </sample>
        }

        public static void Sample_02(IVisio.Application app)
        {
            // <sample id="logging#02" doc="logging.md" line="26">
            var path = VisioAutomation.Application.Logging.LoggingHelper.GetXmlErrorLogFilename(app);
            var log  = new VisioAutomation.Application.Logging.XmlErrorLog(path);

            foreach (var session in log.LogSessions)
            {
                System.Console.WriteLine($"Session at {session.StartTime} ({session.Source}):");
                foreach (var record in session.LogRecords)
                {
                    System.Console.WriteLine($"  {record.Type}/{record.SubType}: {record.Description}");
                }
            }
            // </sample>
        }
    }
}
