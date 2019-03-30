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

        public IDictionary<IVisio.Shape, CustomPropertyDictionary> GetCustomProperties(Models.TargetShapes targets)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            var dicof_shape_to_cpdic = new Dictionary<IVisio.Shape, CustomPropertyDictionary>();
            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return dicof_shape_to_cpdic;
            }

            var listof_cpdic = CustomPropertyHelper.GetCellsAsDictionary(cmdtarget.ActivePage, targets.Shapes, CellValueType.Formula);


            for (int i = 0; i < targets.Shapes.Count; i++)
            {
                var shape = targets.Shapes[i];
                var cpdic = listof_cpdic[i];
                dicof_shape_to_cpdic[shape] = cpdic;
            }

            return dicof_shape_to_cpdic;
        }

        public List<bool> ContainCustomPropertyWithName(Models.TargetShapes targets, string name)
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

        public void DeleteCustomPropertyWithName(Models.TargetShapes targets, string name)
        {
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

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DeleteCustomPropertyWithName)))
            {
                foreach (var shape in targets.Shapes)
                {
                    CustomPropertyHelper.Delete(shape, name);
                }
            }
        }

        public void SetCustomProperty(Models.TargetShapes  targets, string name, CustomPropertyCells customprop)
        {
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

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(SetCustomProperty)))
            {
                foreach (var shape in targets.Shapes)
                {
                    CustomPropertyHelper.Set(shape, name, customprop);
                }
            }
        }
    }
}