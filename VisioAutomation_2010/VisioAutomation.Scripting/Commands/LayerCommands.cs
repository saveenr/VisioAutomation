using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class LayerCommands : CommandSet
    {
        public LayerCommands(Session session) :
            base(session)
        {

        }

        public IVisio.Layer GetLayer(string layername)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            if (layername == null)
            {
                throw new System.ArgumentNullException("layername");
            }

            if (layername.Length < 1)
            {
                throw new System.ArgumentException("layername");
            }

            var application = this.Session.VisioApplication;
            var page = application.ActivePage;
            IVisio.Layer layer = null;
            try
            {
                var layers = page.Layers;
                layer = layers.ItemU[layername];
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                string msg = string.Format("No such layer \"{0}\"", layername);
                throw new VA.Scripting.ScriptingException(msg);
            }
            return layer;
        }

        public IList<IVisio.Layer> GetLayers()
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            var application = this.Session.VisioApplication;
            var page = application.ActivePage;
            return page.Layers.AsEnumerable().ToList();
        }
    }
}