using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class LayerCommands : CommandSet
    {
        internal LayerCommands(Client client) :
            base(client)
        {

        }

        public IVisio.Layer Get(string layername)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            if (layername == null)
            {
                throw new System.ArgumentNullException(nameof(layername));
            }

            if (layername.Length < 1)
            {
                throw new System.ArgumentException("Layer name cannot be empty", nameof(layername));
            }

            var application = cmdtarget.Application;
            var page = application.ActivePage;
            IVisio.Layer layer = null;
            try
            {
                this._client.Output.WriteVerbose("Trying to find layer named \"{0}\"",layername);
                var layers = page.Layers;
                layer = layers.ItemU[layername];
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                string msg = string.Format("No layer with name \"{0}\"", layername);
                throw new VisioAutomation.Exceptions.VisioOperationException(msg);
            }
            return layer;
        }

        public List<IVisio.Layer> Get()
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            var application = cmdtarget.Application;
            var page = application.ActivePage;

            return page.Layers.ToList();
        }
    }
}