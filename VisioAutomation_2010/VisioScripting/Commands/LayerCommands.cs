using System.Collections.Generic;
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

        public IVisio.Layer FindLayersOnActivePageByName(string name)
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);

            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            if (name.Length < 1)
            {
                throw new System.ArgumentException("Layer name cannot be empty", nameof(name));
            }

            var page = cmdtarget.ActivePage;
            IVisio.Layer layer = null;
            try
            {
                this._client.Output.WriteVerbose("Trying to find layer named \"{0}\"",name);
                var layers = page.Layers;
                layer = layers.ItemU[name];
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                string msg = string.Format("No layer with name \"{0}\"", name);
                throw new VisioAutomation.Exceptions.VisioOperationException(msg);
            }
            return layer;
        }

        public List<IVisio.Layer> GetLayersOnActivePage()
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);

            return cmdtarget.ActivePage.Layers.ToList();
        }
    }
}