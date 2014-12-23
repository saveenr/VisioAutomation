using System.Collections.Generic;
using VisioAutomation.Extensions;
using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Documents
{
    public static class DocumentHelper
    {
        /// <summary>
        /// Closes all the documents even if there are unsaved changes
        /// </summary>
        /// <param name="docs"> the Documents object</param>
        public static void ForceCloseAll(IVisio.Documents docs)
        {
            if (docs == null)
            {
                throw new System.ArgumentNullException("docs");
            }

            var application = docs.Application;

            while (docs.Count > 0)
            {
                var active_document = application.ActiveDocument;
                active_document.Close(true);
            }
        }

        public static IVisio.Document OpenStencil(IVisio.Documents docs, string filename)
        {
            var stencil = TryOpenStencil(docs, filename);
            if (stencil == null)
            {
                string msg = string.Format("Could not open stencil \"{0}\"",filename);
                throw new VA.AutomationException(msg);
            }
            return stencil;
        }

        private static IVisio.Document TryOpenStencil(IVisio.Documents docs, string filename)
        {
            const short flags = (short)IVisio.VisOpenSaveArgs.visOpenRO | (short)IVisio.VisOpenSaveArgs.visOpenDocked;
            try
            {
                var doc = docs.OpenEx(filename, flags);
                return doc;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return null;
            }
        }

        public static void Activate(IVisio.Document doc)
        {
            var app = doc.Application;
            var cur_active_doc = app.ActiveDocument;

            // if the doc is already active do nothing
            if (doc == cur_active_doc)
            {
                // do nothing
                return;
            }

            // go through each window and check if it is assigned
            // to the target document
            var appwindows = app.Windows;
            var allwindows = appwindows.AsEnumerable();
            foreach (var curwin in allwindows)
            {
                if (curwin.Document == doc)
                {
                    // we did find one, so activate that window
                    // and then exit the method
                    curwin.Activate();
                    if (app.ActiveDocument != doc)
                    {
                        throw new AutomationException("failed to activate document");
                    }
                    return;
                }
            }

            // If we get here, we couldn't find any matching window
            throw new VA.AutomationException("could not find window for document");
        }

        public static void Close(IVisio.Document doc, bool force_close)
        {
            if (force_close)
            {
                var new_alert_response = VA.Application.AlertResponseCode.No;
                var app = doc.Application;

                using (var alertresponse = new VA.Application.AlertResponseScope(app,new_alert_response))
                {
                    doc.Close();
                }
            }
            else
            {
                doc.Close();
            }
        }
    }
}