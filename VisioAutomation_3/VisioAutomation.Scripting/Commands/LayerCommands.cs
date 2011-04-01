using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;


namespace VisioAutomation.Scripting.Commands
{
    public class LayerCommands : SessionCommands
    {
        public LayerCommands(Session session) :
            base(session)
        {

        }

        public IVisio.Layer GetLayer(string layername)
        {
            if (layername == null)
            {
                throw new ArgumentNullException("layername");
            }

            if (layername.Length < 1)
            {
                throw new ArgumentException("layername");
            }

            var application = this.Session.VisioApplication;
            var page = application.ActivePage;
            IVisio.Layer layer = null;
            try
            {
                var layers = page.Layers;
                layer = layers.ItemU[layername];
            }
            catch (COMException)
            {
                string msg = String.Format("No such layer \"{0}\"", layername);
                throw new AutomationException(msg);
            }
            return layer;
        }

        public IList<IVisio.Layer> GetLayers()
        {
            if (!this.Session.HasActiveDrawing())
            {
                new List<IVisio.Layer>(0);
            }

            var application = this.Session.VisioApplication;
            var page = application.ActivePage;
            return page.Layers.AsEnumerable().ToList();
        }

    }
}