using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class UserDefinedCellCommands : SessionCommands
    {
        public UserDefinedCellCommands(Session session) :
            base(session)
        {

        }
        
        public IDictionary<IVisio.Shape, IList<VA.UserDefinedCells.UserDefinedCell>> GetUserDefinedCells()
        {
            var prop_dic = new Dictionary<IVisio.Shape, IList<VA.UserDefinedCells.UserDefinedCell>>();
            if (!this.Session.HasSelectedShapes())
            {
                return prop_dic;
            }

            var shapes = this.Session.Selection.EnumSelectedShapes().ToList();
            var application = this.Session.Application;
            var page = application.ActivePage;
            var list_user_props = VA.UserDefinedCells.UserDefinedCellsHelper.GetUserDefinedCells(page, shapes);

            for (int i = 0; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                var props = list_user_props[i];
                prop_dic[shape] = props;
            }

            return prop_dic;
        }

        public IList<bool> HasUserDefinedCell(string name)
        {
            if (name == null)
            {
                throw new ArgumentNullException("name");
            }

            if (!this.Session.HasSelectedShapes())
            {
                return new List<bool>();
            }

            var results = (from s in this.Session.Selection.EnumSelectedShapes().ToList()
                           select VA.UserDefinedCells.UserDefinedCellsHelper.HasUserDefinedCell(s, name))
                .ToList();

            return results;
        }

        public void DeleteUserDefinedCell(string name)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            if (name == null)
            {
                throw new ArgumentNullException("name");
            }

            if (name.Length < 1)
            {
                throw new ArgumentException("name");
            }

            var shapes = this.Session.Selection.EnumSelectedShapes().ToList();

            var application = this.Session.Application;
            using (var undoscope = application.CreateUndoScope())
            {
                foreach (var shape in shapes)
                {
                    VA.UserDefinedCells.UserDefinedCellsHelper.DeleteUserDefinedCell(shape, name);
                }
            }
        }

        public void SetUserDefinedCell(VA.UserDefinedCells.UserDefinedCell userprop)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            if (userprop == null)
            {
                throw new ArgumentNullException("userprop");
            }

            var shapes = this.Session.Selection.EnumSelectedShapes().ToList();

            var application = this.Session.Application;
            using (var undoscope = application.CreateUndoScope())
            {
                foreach (var shape in shapes)
                {
                    VA.UserDefinedCells.UserDefinedCellsHelper.SetUserDefinedCell(shape, userprop.Name, userprop.Value, userprop.Prompt);
                }
            }
        }

    }
}