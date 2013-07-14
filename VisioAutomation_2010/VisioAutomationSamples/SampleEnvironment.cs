using System.Linq;
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
                if (app== null)
                {
                    // there is no application object associated with
                    // this session, so create one
                    create_new_app_instance();
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
                        var app_version = app.Version;
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        // If a COMException is thrown, this indicates that the
                        // application object is invalid, so create a new one
                        create_new_app_instance();
                    }                   
                }
                return app;
            }
        }

        private static void create_new_app_instance()
        {
            app = new IVisio.Application();
            var documents = app.Documents;
            documents.Add("");
        }

        public static void SetPageSize(IVisio.Page page, VA.Drawing.Size size)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            var page_sheet = page.PageSheet;

            var update = new VA.ShapeSheet.Update(2);
            update.SetFormula(VA.ShapeSheet.SRCConstants.PageWidth, size.Width);
            update.SetFormula(VA.ShapeSheet.SRCConstants.PageHeight, size.Height);
            update.Execute(page_sheet);
        }

        public static VA.Drawing.Size GetPageSize(IVisio.Page page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            var query = new VA.ShapeSheet.Query.CellQuery();
            var col_height = query.AddColumn(VA.ShapeSheet.SRCConstants.PageHeight);
            var col_width = query.AddColumn(VA.ShapeSheet.SRCConstants.PageWidth);
            var results = query.GetResults<double>(page.PageSheet);
            double height = results.Cells[col_height];
            double width = results.Cells[col_width];
            var s = new VA.Drawing.Size(width, height);
            return s;
        }
    }
}