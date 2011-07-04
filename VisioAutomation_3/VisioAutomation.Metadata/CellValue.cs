using System.Collections.Generic;

namespace VisioAutomation.Metadata
{
    public class CellValue
    {
        public string ID { get; set; }
        public string Enum { get; set; }
        public string Name { get; set; }
        public string Value { get; set; }
        public string AutomationConstant { get; set; }
    }

    public class CellValueEnum
    {
        public string Name { get; set; }
        public string[] CellNameCodes { get; set; }
        public List<CellValue> Items { get; set; }
    }
}