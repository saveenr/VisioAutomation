using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.UserDefinedCells
{
    public class UserDefinedCell
    {
        public string Name { get; set; }
        public string Value { get; set; }
        public string Prompt { get; set; }

        internal static readonly UserDefinedCellQuery query = new UserDefinedCellQuery();

        public UserDefinedCell(string name)
        {
            UserDefinedCellsHelper.CheckValidName(name);
            this.Name = name;
        }

        public UserDefinedCell(string name, string value)
        {
            UserDefinedCellsHelper.CheckValidName(name);

            if (value == null)
            {
                throw new System.ArgumentNullException("value");
            }

            this.Name = name;
            this.Value = value;
        }

        public UserDefinedCell(string name, string value, string prompt)
        {
            UserDefinedCellsHelper.CheckValidName(name);

            if (value == null)
            {
                throw new System.ArgumentNullException("value");
            }
            
            this.Name = name;
            this.Value = value;
            this.Prompt = prompt;
        }

        public override string ToString()
        {
            string s = string.Format("(Name={0},Value={1},Prompt={2})",
                                     this.Name,
                                     this.Value,
                                     this.Prompt);
            return s;
        }

        internal class UserDefinedCellQuery : VA.ShapeSheet.Query.SectionQuery
        {
            public VA.ShapeSheet.Query.QueryColumn Value { get; set; }
            public VA.ShapeSheet.Query.QueryColumn Prompt { get; set; }

            public UserDefinedCellQuery() :
                base(IVisio.VisSectionIndices.visSectionUser)
            {
                Value = this.AddColumn(VA.ShapeSheet.SRCConstants.User_Value, "Value");
                Prompt = this.AddColumn(VA.ShapeSheet.SRCConstants.User_Prompt, "Prompt");
            }
        }

    }
}