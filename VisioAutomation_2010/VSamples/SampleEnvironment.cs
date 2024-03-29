﻿using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VSamples
{
    public class SampleEnvironment
    {
        private static IVisio.Application _app;

        public static IVisio.Application Application
        {
            get
            {
                if (SampleEnvironment._app == null)
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
                        var app_version = SampleEnvironment._app.ProductName;
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        // If a COMException is thrown, this indicates that the
                        // application object is invalid, so create a new one
                        SampleEnvironment.create_new_app_instance();
                    }
                }

                return SampleEnvironment._app;
            }
        }

        private static void create_new_app_instance()
        {
            SampleEnvironment._app = new IVisio.Application();
            var documents = SampleEnvironment._app.Documents;
            documents.Add("");
        }

        public static void SetPageSize(IVisio.Page page, VA.Core.Size size)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            var page_sheet = page.PageSheet;

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetValue(VA.Core.SrcConstants.PageWidth, size.Width);
            writer.SetValue(VA.Core.SrcConstants.PageHeight, size.Height);

            writer.Commit(page_sheet, VisioAutomation.Core.CellValueType.Formula);
        }

        public static VA.Core.Size GetPageSize(IVisio.Page page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            var query = new CellQuery();
            var col_height = query.Columns.Add(VA.Core.SrcConstants.PageHeight);
            var col_width = query.Columns.Add(VA.Core.SrcConstants.PageWidth);

            var cellqueryresults = query.GetResults<double>(page.PageSheet);
            var row = cellqueryresults[0];
            double height = row[col_height];
            double width = row[col_width];
            var s = new VA.Core.Size(width, height);
            return s;
        }
    }
}