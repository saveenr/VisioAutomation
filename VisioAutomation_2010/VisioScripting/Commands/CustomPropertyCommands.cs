using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Shapes;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class CustomPropertyCommands : CommandSet
    {
        internal CustomPropertyCommands(Client client) :
            base(client)
        {

        }

        public IDictionary<IVisio.Shape, CustomPropertyDictionary> GetCustomPropertiesAsShapeDictionary(TargetShapes targetshapes, VisioAutomation.Core.CellValueType type)
        {
            targetshapes = targetshapes.ResolveToShapes(this._client);
            var dicof_shape_to_cpdic = new Dictionary<IVisio.Shape, CustomPropertyDictionary>();
            var listof_cpdic = GetCustomProperties(targetshapes, type);

            for (int i = 0; i < listof_cpdic.Count; i++)
            {
                var shape = targetshapes.Shapes[i];
                var cpdic = listof_cpdic[i];
                dicof_shape_to_cpdic[shape] = cpdic;
            }

            return dicof_shape_to_cpdic;
        }

        public List<CustomPropertyDictionary> GetCustomProperties(TargetShapes targetshapes, VisioAutomation.Core.CellValueType type)
        {
            targetshapes = targetshapes.ResolveToShapes(this._client);

            if (targetshapes.Shapes.Count < 1)
            {
                return new List<CustomPropertyDictionary>(0);
            }

            var page = targetshapes.Shapes[0].ContainingPage;
            var shapeidpairs = targetshapes.ToShapeIDPairs();
            var listof_cpdic = CustomPropertyHelper.GetDictionary(page, shapeidpairs, type);

            return listof_cpdic;
        }

        public List<bool> ContainCustomPropertyWithName(TargetShapes targetshapes, string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            targetshapes = targetshapes.ResolveToShapes(this._client);

            var results = new List<bool>(targetshapes.Shapes.Count);
            var values = targetshapes.Shapes.Select(shape => CustomPropertyHelper.Contains(shape, name));
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

            targetshapes = targetshapes.ResolveToShapes(this._client);

            if (targetshapes.Shapes.Count < 1)
            {
                return;
            }

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DeleteCustomPropertyWithName)))
            {
                foreach (var shape in targetshapes.Shapes)
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

            targetshapes = targetshapes.ResolveToShapes(this._client);

            if (targetshapes.Shapes.Count < 1)
            {
                return;
            }

            customprop.EncodeValues();

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(SetCustomProperty)))
            {
                foreach (var shape in targetshapes.Shapes)
                {
                    CustomPropertyHelper.Set(shape, name, customprop);
                }
            }
        }
    }
}