using System.Collections.Generic;
using System.Linq;
using VA_UDC = VisioAutomation.Shapes.UserDefinedCells;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.Commands
{
    public class UserDefinedCellCommands : CommandSet
    {
        internal UserDefinedCellCommands(Client client) :
            base(client)
        {

        }

        public IDictionary<IVisio.Shape, IList<VA_UDC.UserDefinedCell>> Get(IList<IVisio.Shape> target_shapes)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var prop_dic = new Dictionary<IVisio.Shape, IList<VA_UDC.UserDefinedCell>>();

            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return prop_dic;
            }

            var application = this.Client.Application.Get();
            var page = application.ActivePage;
            var list_user_props = VA_UDC.UserDefinedCellsHelper.Get(page, shapes);

            for (int i = 0; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                var props = list_user_props[i];
                prop_dic[shape] = props;
            }

            return prop_dic;
        }

        public IList<bool> Contains(IList<IVisio.Shape> target_shapes, string name)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return new List<bool>();
            }

            var all_shapes = this.Client.Selection.GetShapes();
            var results = all_shapes.Select(s => VA_UDC.UserDefinedCellsHelper.Contains(s, name)).ToList();

            return results;
        }
       
        public void Delete(IList<IVisio.Shape> target_shapes, string name)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            } 

            if (name == null)
            {
                throw new System.ArgumentNullException("name cannot be null","name");
            }

            if (name.Length < 1)
            {
                throw new System.ArgumentException("name cannot be empty", nameof(name));
            }

            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Delete User-Defined Cell"))
            {
                foreach (var shape in shapes)
                {
                    VA_UDC.UserDefinedCellsHelper.Delete(shape, name);
                }
            }
        }

        public void Set(IList<IVisio.Shape> target_shapes, VA_UDC.UserDefinedCell userdefinedcell)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            }

            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Set User-Defined Cell"))
            {
                foreach (var shape in shapes)
                {
                    VA_UDC.UserDefinedCellsHelper.Set(shape, userdefinedcell.Name, userdefinedcell.Value.Formula.Value, userdefinedcell.Prompt.Formula.Value);
                }
            }
        }
    }
}