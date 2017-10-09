using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Shapes;
using VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class CustomPropertyCommands : CommandSet
    {
        internal CustomPropertyCommands(Client client) :
            base(client)
        {

        }

        public IDictionary<IVisio.Shape, CustomPropertyDictionary> Get(VisioScripting.Models.TargetShapes targets)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);

            var prop_dic = new Dictionary<IVisio.Shape, CustomPropertyDictionary>();
            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return prop_dic;
            }

            var page = cmdtarget.Page;

            var list_custom_props = CustomPropertyHelper.GetCells(page, targets.Shapes, CellValueType.Formula);

            for (int i = 0; i < targets.Shapes.Count; i++)
            {
                var shape = targets.Shapes[i];
                var props = list_custom_props[i];
                prop_dic[shape] = props;
            }

            return prop_dic;
        }

        public List<bool> Contains(VisioScripting.Models.TargetShapes targets, string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            targets = targets.ResolveShapes(this._client);

            var results = new List<bool>(targets.Shapes.Count);
            var values = targets.Shapes.Select(shape => CustomPropertyHelper.Contains(shape, name));
            results.AddRange(values);

            return results;
        }

        public void Delete(VisioScripting.Models.TargetShapes targets, string name)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);

            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            if (name.Length < 1)
            {
                throw new System.ArgumentException("name cannot be empty", nameof(name));
            }

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return;
            }

            using (var undoscope = this._client.Application.NewUndoScope("Delete Custom Property"))
            {
                foreach (var shape in targets.Shapes)
                {
                    CustomPropertyHelper.Delete(shape, name);
                }
            }
        }

        public void Set(VisioScripting.Models.TargetShapes  targets, string name, CustomPropertyCells customprop)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            if (customprop == null)
            {
                throw new System.ArgumentNullException(nameof(customprop));
            }

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return;
            }

            customprop.EncodeValues();

            using (var undoscope = this._client.Application.NewUndoScope("Set Custom Property"))
            {
                foreach (var shape in targets.Shapes)
                {
                    CustomPropertyHelper.Set(shape, name, customprop);
                }
            }
        }
    }
}