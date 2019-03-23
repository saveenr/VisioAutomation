﻿using VisioAutomation.ShapeSheet.Query;
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

        public static void SetPageSize(IVisio.Page page, VA.Geometry.Size size)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            var page_sheet = page.PageSheet;

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetValue(VA.ShapeSheet.SrcConstants.PageWidth, size.Width);
            writer.SetValue(VA.ShapeSheet.SrcConstants.PageHeight, size.Height);

            writer.CommitFormulas(page_sheet);
        }

        public static VA.Geometry.Size GetPageSize(IVisio.Page page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            var query = new CellQuery();
            var col_height = query.Columns.Add(VA.ShapeSheet.SrcConstants.PageHeight,nameof(VA.ShapeSheet.SrcConstants.PageHeight));
            var col_width = query.Columns.Add(VA.ShapeSheet.SrcConstants.PageWidth, nameof(VA.ShapeSheet.SrcConstants.PageWidth));

            var cellqueryresults = query.GetResults<double>(page.PageSheet);
            var row = cellqueryresults[0];
            double height = row[col_height];
            double width = row[col_width];
            var s = new VA.Geometry.Size(width, height);
            return s;
        }
    }
}