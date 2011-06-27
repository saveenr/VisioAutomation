using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Data;
using ExcelUtil;

namespace VisioAutomation.Metadata
{


    public class Cell
    {
        public string ID;
        public string Name;
        public string NameFormatString;
        public string Object;
        public string NameType;
        public string DataType;
        public string ContentType;
        public string Unit;
        public string SectionIndex;
        public string RowIndex;
        public string MinVersion;
        public string MaxVersion;
        public string CellIndex;
        public string MSDN;
    }

    public class MetadataDB
    {
        ExcelXmlToDataSetConverter converter = new ExcelUtil.ExcelXmlToDataSetConverter();

        public List<Cell> GetCells()
        {
            converter.Parse(VisioAutomation.Metadata.Properties.Resources.cells);
            var cells_table = converter.DataSet.Tables[0];

            var cells = new List<Cell>();
            foreach (var item in cells_table.AsEnumerable())
            {
                var c = new Cell();
                cells.Add(c);
                c.ID = item.Field<string>("ID");
                c.Name = item.Field<string>("Name");
                c.NameFormatString = item.Field<string>("NameFormatString");
                c.Object = item.Field<string>("Object");
                c.NameType = item.Field<string>("NameType");
                c.DataType = item.Field<string>("DataType");
                c.ContentType = item.Field<string>("ContentType");
                c.Unit = item.Field<string>("Unit");
                c.SectionIndex = item.Field<string>("SectionIndex");
                c.RowIndex = item.Field<string>("RowIndex");
                c.MinVersion = item.Field<string>("MinVersion");
                c.MaxVersion = item.Field<string>("MaxVersion");
                c.CellIndex = item.Field<string>("CellIndex");
                c.MSDN = item.Field<string>("MSDN");

            }

            return cells;
        }

    
        public List<Section> GetSections()
        {
            converter.Parse(VisioAutomation.Metadata.Properties.Resources.sections);
            var sections_table = converter.DataSet.Tables[0];
            
            var sections = new List<Section>();
            foreach (var item in sections_table.AsEnumerable())
            {
                var c = new Section();
                sections.Add(c);
                c.ID = item.Field<string>("ID");
                c.DisplayName = item.Field<string>("DisplayName");
                c.Name = item.Field<string>("Name");
            }
            
            return sections;
        }

        public List<AutomationConstant> GetAutomationEnums()
        {

            converter.Parse(VisioAutomation.Metadata.Properties.Resources.automationconstants);
            var automationenums_table = converter.DataSet.Tables[0];

            var constants = new List<AutomationConstant>();
            foreach (var item in automationenums_table.AsEnumerable())
            {
                var c = new AutomationConstant();
                constants.Add(c);
                c.ID = item.Field<string>("ID");
                c.Enum = item.Field<string>("EnumName");
                c.Name = item.Field<string>("ValueName");
                c.Value = int.Parse(item.Field<string>("ValueInt"));
            }
            return constants;
        }

        public List<CellValue> GetCellValues()
        {
            converter.Parse(VisioAutomation.Metadata.Properties.Resources.cellvalues);
            var cellvalues_table = converter.DataSet.Tables[0];
            var constants = new List<CellValue>();
            foreach (var item in cellvalues_table.AsEnumerable())
            {
                var c = new CellValue();
                constants.Add(c);
                c.ID = item.Field<string>("ID");
                c.Enum = item.Field<string>("Enum");
                c.Name = item.Field<string>("Name");

                bool s;
                int v;
                s = int.TryParse(item.Field<string>("Value"), out v);
                if (s)
                {
                    c.Value = v;
                }
                else
                {
                    c.Value = null;
                }

                c.AutomationConstant = item.Field<string>("AutomationConstant");

            }
            return constants;
        }


    }

    public class CellValue
    {
        public string ID;
        public string Enum;
        public string Name;
        public int? Value;
        public string AutomationConstant;
    }

    public class AutomationConstant
    {
        public string ID;
        public string Enum;
        public string Name;
        public int Value;
    }

    public class Section
    {
        public string ID;
        public string Name;
        public string DisplayName;
        
    }
}
