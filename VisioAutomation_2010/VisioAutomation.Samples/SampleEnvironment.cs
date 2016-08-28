using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomationSamples
{
    public class SampleEnvironment
    {
        private static IVisio.Application app;

        public static IVisio.Application Application
        {
            get
            {
                if (SampleEnvironment.app== null)
                {
                    // there is no application object associated with
                    // this session, so create one
                    SampleEnvironment.create_new_app_instance();
                }
                else
                {
                    // there is an application object associated with this session

                    // before we continue we should try to validate that the
                    // application is valid - the user might have closed the application
                    // leaving us with an application object that is invalid

                    try
                    {
                        // try to do something simple, read-only, and fast with the application object
                        var app_version = SampleEnvironment.app.ProductName;
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        // If a COMException is thrown, this indicates that the
                        // application object is invalid, so create a new one
                        SampleEnvironment.create_new_app_instance();
                    }                   
                }
                return SampleEnvironment.app;
            }
        }

        private static void create_new_app_instance()
        {
            SampleEnvironment.app = new IVisio.Application();
            var documents = SampleEnvironment.app.Documents;
            documents.Add("");
        }

        public static void SetPageSize(IVisio.Page page, VA.Drawing.Size size)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            var page_sheet = page.PageSheet;

            var writer = new FormulaWriterSRC(2);
            writer.SetFormula(VA.ShapeSheet.SRCConstants.PageWidth, size.Width);
            writer.SetFormula(VA.ShapeSheet.SRCConstants.PageHeight, size.Height);
            writer.Commit(page_sheet);
        }

        public static VA.Drawing.Size GetPageSize(IVisio.Page page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            var query = new VisioAutomation.ShapeSheet.Queries.Query();
            var col_height = query.AddCell(VA.ShapeSheet.SRCConstants.PageHeight,"PageHeight");
            var col_width = query.AddCell(VA.ShapeSheet.SRCConstants.PageWidth, "PageWidth");

            var ss = new ShapeSheetSurface(page.PageSheet);
            var results = query.GetResults<double>(ss);
            double height = results.Cells[col_height];
            double width = results.Cells[col_width];
            var s = new VA.Drawing.Size(width, height);
            return s;
        }
    }
}