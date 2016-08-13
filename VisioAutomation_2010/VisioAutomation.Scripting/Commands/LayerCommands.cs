using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.Commands
{
    public class LayerCommands : CommandSet
    {
        internal LayerCommands(Client client) :
            base(client)
        {

        }

        public IVisio.Layer Get(string layername)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            if (layername == null)
            {
                throw new System.ArgumentNullException(nameof(layername));
            }

            if (layername.Length < 1)
            {
                throw new System.ArgumentException("Layer name cannot be empty", nameof(layername));
            }

            var application = this._client.Application.Get();
            var page = application.ActivePage;
            IVisio.Layer layer = null;
            try
            {
                this._client.WriteVerbose("Trying to find layer named \"{0}\"",layername);
                var layers = page.Layers;
                layer = layers.ItemU[layername];
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                string msg = string.Format("No layer with name \"{0}\"", layername);
                throw new VisioOperationException(msg);
            }
            return layer;
        }

        public IList<IVisio.Layer> Get()
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var application = this._client.Application.Get();
            var page = application.ActivePage;
            return page.Layers.ToEnumerable().ToList();
        }
    }
}