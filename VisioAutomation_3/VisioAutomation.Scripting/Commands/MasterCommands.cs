using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class MasterCommands : SessionCommands
    {
        public MasterCommands(Session session) :
            base(session)
        {

        }

        public IList<IVisio.Master> GetMasters()
        {
            if (!this.Session.HasActiveDrawing)
            {
                this.Session.Write(OutputStream.Verbose,"No Active Document - 0 Masters");
                new List<IVisio.Master>(0);
            }

            var application = this.Session.VisioApplication;
            var doc = application.ActiveDocument;
            var doc_masters = doc.Masters;
            var masters = doc_masters.AsEnumerable().ToList();
            return masters;
        }

        public IList<IVisio.Master> GetMasters(IVisio.Document doc)
        {
            var doc_masters = doc.Masters;
            var masters = doc_masters.AsEnumerable().ToList();
            return masters;
        }

        public IVisio.Master GetMaster(string name)
        {
            if (name == null)
            {
                throw new ArgumentNullException("name");
            }

            if (name.Length < 1)
            {
                throw new ArgumentException("mastername");
            }


            IVisio.Master master = null;
            try
            {
                var application = this.Session.VisioApplication;
                var active_document = application.ActiveDocument;
                var masters = active_document.Masters;
                master = masters.ItemU[name];
            }
            catch (COMException)
            {
                string msg = String.Format("No such master \"{0}\"", name);
                throw new AutomationException(msg);
            }
            return master;
        }

        public IVisio.Master GetMaster(string master, string stencil)
        {
            if (master == null)
            {
                throw new ArgumentNullException("master");
            }

            if (stencil == null)
            {
                throw new ArgumentNullException("stencil");
            }

            IVisio.Document stencil_doc = null;
            try
            {
                var application = this.Session.VisioApplication;
                var documents = application.Documents;
                stencil_doc = documents[stencil];
            }
            catch (COMException)
            {
                string msg = String.Format("No such stencil \"{0}\"", stencil);
                throw new AutomationException(msg);
            }


            IVisio.Master masterobj = null;
            try
            {
                var masters = stencil_doc.Masters;
                masterobj = masters.ItemU[master];
            }
            catch (COMException)
            {
                string msg = String.Format("No such master \"{0}\" in \"{1}\"", master, stencil);
                throw new AutomationException(msg);
            }

            return masterobj;
        }

        public IVisio.Shape DropMaster(IVisio.Master master, double x, double y)
        {
            var application = this.Session.VisioApplication;
            var page = application.ActivePage;
            var shape = page.Drop(master, x, y);
            return shape;
        }

        public short[] DropMasters(IList<IVisio.Master> masters, IList<VA.Drawing.Point> points)
        {
            if (masters == null)
            {
                throw new ArgumentNullException("points");
            }

            if (points == null)
            {
                throw new ArgumentNullException("points");
            }

            if (!this.Session.HasActiveDrawing)
            {
                throw new AutomationException("No active page");
            }

            var application = this.Session.VisioApplication;
            var page = application.ActivePage;
            var shapeids = page.DropManyU(masters, points);
            return shapeids;
        }

        public IVisio.Master NewMaster(IVisio.Document stencil, string name)
        {

            var masters = stencil.Masters;

            var master = masters.AddEx(IVisio.VisMasterTypes.visTypeMaster);
            master.Name = name;

            return master;
        }
    }
}