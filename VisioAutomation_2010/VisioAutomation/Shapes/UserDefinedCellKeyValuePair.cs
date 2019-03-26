using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class UserDefinedCellKeyValuePair
    {
        public readonly string Name;
        public readonly UserDefinedCellCells Cells;

        public UserDefinedCellKeyValuePair(string name, UserDefinedCellCells cells)
        {
            this.Name = name;
            this.Cells = cells;
        }
    }
}