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

        public IDictionary<IVisio.Shape, CustomPropertyDictionary> GetCustomProperties(TargetShapes targetshapes)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            var dicof_shape_to_cpdic = new Dictionary<IVisio.Shape, CustomPropertyDictionary>();
            targetshapes = targetshapes.Resolve(this._client);

            if (targetshapes.Items.Count < 1)
            {
                return dicof_shape_to_cpdic;
            }

            var shapeidpairs = targetshapes.ToShapeIDPairs();
            var listof_cpdic = CustomPropertyHelper.GetCellsAsDictionary(cmdtarget.ActivePage, shapeidpairs, CellValueType.Formula);


            for (int i = 0; i < targetshapes.Items.Count; i++)
            {
                var shape = targetshapes.Items[i];
                var cpdic = listof_cpdic[i];
                dicof_shape_to_cpdic[shape] = cpdic;
            }

            return dicof_shape_to_cpdic;
        }

        public List<bool> ContainCustomPropertyWithName(TargetShapes targetshapes, string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            targetshapes = targetshapes.Resolve(this._client);

            var results = new List<bool>(targetshapes.Items.Count);
            var values = targetshapes.Items.Select(shape => CustomPropertyHelper.Contains(shape, name));
            results.AddRange(values);

            return results;
        }

        public void DeleteCustomPropertyWithName(TargetShapes targetshapes, string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            if (name.Length < 1)
            {
                throw new System.ArgumentException("name cannot be empty", nameof(name));
            }

            targetshapes = targetshapes.Resolve(this._client);

            if (targetshapes.Items.Count < 1)
            {
                return;
            }

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DeleteCustomPropertyWithName)))
            {
                foreach (var shape in targetshapes.Items)
                {
                    CustomPropertyHelper.Delete(shape, name);
                }
            }
        }

        public void SetCustomProperty(TargetShapes targetshapes, string name, CustomPropertyCells customprop)
        {
            if (customprop == null)
            {
                throw new System.ArgumentNullException(nameof(customprop));
            }

            targetshapes = targetshapes.Resolve(this._client);

            if (targetshapes.Items.Count < 1)
            {
                return;
            }

            customprop.EncodeValues();

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(SetCustomProperty)))
            {
                foreach (var shape in targetshapes.Items)
                {
                    CustomPropertyHelper.Set(shape, name, customprop);
                }
            }
        }
    }
}