using System.Collections.Generic;
using System.Linq;
using VACUSTPROP = VisioAutomation.Shapes.CustomProperties;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.Commands
{
    public class CustomPropCommands : CommandSet
    {
        internal CustomPropCommands(Client client) :
            base(client)
        {

        }

        public IDictionary<IVisio.Shape, VACUSTPROP.CustomPropertyDictionary> Get(TargetShapes targets)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var prop_dic = new Dictionary<IVisio.Shape, VACUSTPROP.CustomPropertyDictionary>();
            var shapes = targets.ResolveShapesEx(this._client);

            if (shapes.Shapes.Count < 1)
            {
                return prop_dic;
            }

            var application = this._client.Application.Get();
            var page = application.ActivePage;

            var list_custom_props = VACUSTPROP.CustomPropertyHelper.Get(page, shapes.Shapes);

            for (int i = 0; i < shapes.Shapes.Count; i++)
            {
                var shape = shapes.Shapes[i];
                var props = list_custom_props[i];
                prop_dic[shape] = props;
            }

            return prop_dic;
        }

        public List<bool> Contains(TargetShapes targets, string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            var shapes = targets.ResolveShapesEx(this._client);

            var results = new List<bool>(shapes.Shapes.Count);
            foreach (var shape in shapes.Shapes)
            {
                results.Add(VACUSTPROP.CustomPropertyHelper.Contains(shape, name));
            }

            return results;
        }

        public void Delete(TargetShapes targets, string name)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();
            
            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            if (name.Length < 1)
            {
                throw new System.ArgumentException("name cannot be empty", nameof(name));
            }

            var shapes = targets.ResolveShapesEx(this._client);

            if (shapes.Shapes.Count < 1)
            {
                return;
            }

            using (var undoscope = this._client.Application.NewUndoScope("Delete Custom Property"))
            {
                foreach (var shape in shapes.Shapes)
                {
                    VACUSTPROP.CustomPropertyHelper.Delete(shape, name);
                }
            }
        }

        public void Set(TargetShapes  targets, string name, VACUSTPROP.CustomPropertyCells customprop)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();
            
            if (customprop == null)
            {
                throw new System.ArgumentNullException(nameof(customprop));
            }

            var shapes = targets.ResolveShapesEx(this._client);

            if (shapes.Shapes.Count < 1)
            {
                return;
            }

            using (var undoscope = this._client.Application.NewUndoScope("Set Custom Property"))
            {
                foreach (var shape in shapes.Shapes)
                {
                    VACUSTPROP.CustomPropertyHelper.Set(shape, name, customprop);
                }
            }
        }

        public IEnumerable<IVisio.Shape> EnumerateAndSelect(IEnumerable<IVisio.Shape> shapes)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();
            
            if (shapes == null)
            {
                throw new System.ArgumentNullException(nameof(shapes));
            }

            foreach (var shape in shapes)
            {
                this._client.Selection.SelectNone();
                this._client.Selection.Select(shape);
                yield return shape;
            }
        }
    }
}