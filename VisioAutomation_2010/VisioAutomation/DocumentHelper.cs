using System;
using VisioAutomation.Extensions;
using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation
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

            while (docs.Count > 0)
            {
                var application = docs.Application;
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

        public static IVisio.Document TryOpenStencil(IVisio.Documents docs, string filename)
        {
            if (filename == null)
            {
                throw new System.ArgumentNullException("filename");
            }

            short flags = (short)IVisio.VisOpenSaveArgs.visOpenRO | (short)IVisio.VisOpenSaveArgs.visOpenDocked;
            try
            {
                var doc = docs.OpenEx(filename, flags);
                return doc;
            }
            catch (System.Runtime.InteropServices.COMException comexc)
            {
                return null;
            }
        }

        public static void Activate(IVisio.Document doc)
        {
            bool found_window = false;
            IVisio.Window the_window = null;
            var application = doc.Application;
            var appwindows = application.Windows;
            var allwindows = appwindows.AsEnumerable();
            foreach (var curwin in allwindows)
            {
                if (curwin.Document == doc)
                {
                    found_window = true;
                    the_window = curwin;
                    break;
                }
            }

            if (!found_window)
            {
                throw new VA.AutomationException("could not find window for doc");
            }

            if (the_window == null)
            {
                throw new VA.AutomationException("Internal Error");
            }

            the_window.Activate();

            if (application.ActiveDocument != doc)
            {
                throw new AutomationException("failed to activate document");
            }
        }

        public static void Close(IVisio.Document doc, bool force_close)
        {
            if (force_close)
            {
                var new_alert_response = VA.UI.AlertResponseCode.No;
                var app = doc.Application;

                using (var alertresponse = app.CreateAlertResponseScope(new_alert_response))
                {
                    doc.Close();
                }
            }
            else
            {
                doc.Close();
            }
        }

        public static IVisio.Document NewStencil(IVisio.Documents documents)
        {

            var doc = documents.AddEx(string.Empty, IVisio.VisMeasurementSystem.visMSDefault,
                                      (int)IVisio.VisOpenSaveArgs.visAddStencil +
                                      (int)IVisio.VisOpenSaveArgs.visOpenDocked,
                                      0);
            return doc;
        }

        public static IVisio.Document TryGetDocument(IVisio.Documents documents, string name)
        {
            try
            {
                var stencil_doc = documents[name];
                return stencil_doc;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return null;
            }
        }
    }
}