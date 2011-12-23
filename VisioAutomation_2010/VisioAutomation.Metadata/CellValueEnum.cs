using System.Collections.Generic;

namespace VisioAutomation.Metadata
{
    public class CellValueEnum
    {
        public string Name { get; set; }
        public string[] CellNameCodes { get; set; }
        public List<CellValue> Items { get; set; }
    }
}