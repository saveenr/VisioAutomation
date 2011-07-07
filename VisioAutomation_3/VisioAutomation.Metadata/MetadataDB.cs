using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Reflection;
using System.Xml;
using System.Xml.Linq;
using VA = VisioAutomation;

namespace VisioAutomation.Metadata
{
    public class MetadataDB
    {
        private List<Cell> _cells;
        private List<CellValue> _cellvals;
        private List<Section> _sections;
        private List<AutomationEnum> _autoenums;
        private List<CellValueEnum> _cellvalueenum;
        private Dictionary<string, AutomationEnum> _name_to_autoenums;
        private Dictionary<string, CellValueEnum> _name_to_cellvalueenums;
        private List<AutomationConstant> _constants;
        private Dictionary<string, AutomationConstant> _name_to_constants;
        private Dictionary<int, Section> _int_to_section;
        private Dictionary<string, Cell> _namecode_to_cell;

        public void Save(string path)
        {
            var cell_writer = new XmlPersist.XmlTableWriter<Cell>();
            var cv_writer = new XmlPersist.XmlTableWriter<CellValue>();
            var sec_writer = new XmlPersist.XmlTableWriter<Section>();
            var con_writer = new XmlPersist.XmlTableWriter<AutomationConstant>();

            string cells_filename = System.IO.Path.Combine(path, "cells.xml");
            string cv_filename = System.IO.Path.Combine(path, "cellvalues.xml");
            string sec_filename = System.IO.Path.Combine(path, "sections.xml");
            string con_filename = System.IO.Path.Combine(path, "constants.xml");

            cell_writer.SaveToFile(this.Cells, cells_filename);
            cv_writer.SaveToFile(this.CellValues, cv_filename);
            sec_writer.SaveToFile(this.Sections, sec_filename);
            con_writer.SaveToFile(this.Constants, con_filename);
        }

        /*
         * NOTES
         * - Cell Names are not unique - use Cell.NameCode instead
         */

        private MetadataDB()
        {
        }

        private void init_from_doms(XDocument dom_cells, XDocument dom_sec, XDocument dom_constants, XDocument dom_cv)
        {
// Load constants
            var conreader = new XmlPersist.XmlTableReader<AutomationConstant>();

            this._constants = conreader.Load(dom_constants).ToList();
            this._name_to_constants = this._constants.ToDictionary(i => i.Name, i => i);

            // Create enums from the constants
            this._autoenums = new List<AutomationEnum>();
            this._name_to_autoenums = new Dictionary<string, AutomationEnum>();
            foreach (var constant in this.Constants)
            {
                var enum_name = constant.Enum;
                AutomationEnum a_enum = null;
                this._name_to_autoenums.TryGetValue(constant.Enum, out a_enum);
                if (a_enum == null)
                {
                    a_enum = new AutomationEnum(constant.Enum);
                    this._name_to_autoenums[a_enum.Name] = a_enum;
                    this._autoenums.Add(a_enum);
                }

                a_enum.Add(constant.Name, constant.GetValueAsInt());
            }

            // Load cell values
            var cvreader = new XmlPersist.XmlTableReader<CellValue>();
            this._cellvals = cvreader.Load(dom_cv).ToList();

            // Load Cell data
            var cellreader = new XmlPersist.XmlTableReader<Cell>();
            this._cells = cellreader.Load(dom_cells).ToList();

            // Index cells
            this._namecode_to_cell = this.Cells.ToDictionary(i => i.NameCode, i => i);

            // Initialize sections
            var secreader = new XmlPersist.XmlTableReader<Section>();
            this._sections = secreader.Load(dom_sec).ToList();

            this._int_to_section = new Dictionary<int, Section>();
            foreach (var section in this.Sections)
            {
                string secindex_name = section.Enum;
                int secindex_int = this.GetAutomationConstantByName(secindex_name).GetValueAsInt();

                this._int_to_section[secindex_int] = section;
            }

            // load the cell values

            this._name_to_cellvalueenums = new Dictionary<string, CellValueEnum>();

            foreach (var c in this.CellValues)
            {
                string enum_name = c.Enum;

                bool s;
                CellValueEnum cve;
                s = this._name_to_cellvalueenums.TryGetValue(enum_name, out cve);
                if (!s)
                {
                    cve = new CellValueEnum();
                    cve.Name = enum_name;
                    cve.Items = new List<CellValue>();
                    this._name_to_cellvalueenums[enum_name] = cve;
                }

                cve.Items.Add(c);
            }
        }

        public static MetadataDB Load()
        {
            var db = new MetadataDB();
            var dom_constants = System.Xml.Linq.XDocument.Parse(VA.Metadata.Properties.Resources.constants);
            var dom_cells = System.Xml.Linq.XDocument.Parse(VA.Metadata.Properties.Resources.cells);
            var dom_cv = System.Xml.Linq.XDocument.Parse(VA.Metadata.Properties.Resources.cellvalues);
            var dom_sec = System.Xml.Linq.XDocument.Parse(VA.Metadata.Properties.Resources.sections);
            db.init_from_doms(dom_cells, dom_sec, dom_constants, dom_cv);
            return db;
        }

        public static MetadataDB Load(string path)
        {
            string cells_filename = System.IO.Path.Combine(path, "cells.xml");
            string cv_filename = System.IO.Path.Combine(path, "cellvalues.xml");
            string sec_filename = System.IO.Path.Combine(path, "sections.xml");
            string con_filename = System.IO.Path.Combine(path, "constants.xml");

            var cellreader = new XmlPersist.XmlTableReader<Cell>();
            var cvreader = new XmlPersist.XmlTableReader<CellValue>();
            var secreader = new XmlPersist.XmlTableReader<Section>();
            var conreader = new XmlPersist.XmlTableReader<AutomationConstant>();

            var dom_constants = System.Xml.Linq.XDocument.Load(con_filename);
            var dom_cells = System.Xml.Linq.XDocument.Load(cells_filename);
            var dom_cv = System.Xml.Linq.XDocument.Load(cv_filename);
            var dom_sec = System.Xml.Linq.XDocument.Load(sec_filename);

            var db = new MetadataDB();
            db.init_from_doms(dom_cells, dom_sec, dom_constants, dom_cv);
            return db;
        }

        public List<Cell> Cells
        {
            get { return _cells; }
        }

        public List<Section> Sections
        {
            get { return this._sections; }
        }

        public List<AutomationConstant> Constants
        {
            get { return this._constants; }
        }

        public List<AutomationEnum> AutomationEnums
        {
            get { return this._autoenums; }
        }

        public Cell GetCellByNameCode(string name)
        {
            return this._namecode_to_cell[name];
        }

        public AutomationEnum GetAutomationEnumByName(string name)
        {
            return this._name_to_autoenums[name];
        }

        public AutomationConstant GetAutomationConstantByName(string name)
        {
            return this._name_to_constants[name];
        }

        public Section GetSectionBySectionIndex(int sectionindex)
        {
            return this._int_to_section[sectionindex];
        }

        public List<CellValue> CellValues
        {
            get { return this._cellvals; }
        }

        public void ExportCode(string filename)
        {
            var fp = System.IO.File.CreateText(filename);
            var records = new[]
                              {
                                  new { type = typeof (VA.Metadata.Cell), col = this.Cells.Cast<object>().ToList()},
                                  new { type = typeof (VA.Metadata.CellValue), col = this.CellValues.Cast<object>().ToList() },
                                  new { type = typeof (VA.Metadata.Section), col = this.Sections.Cast<object>().ToList() },
                                  new { type = typeof (VA.Metadata.AutomationConstant), col = this.Constants.Cast<object>().ToList() }
                              };

            fp.WriteLine("// Metadata DB Exported on {0}", System.DateTime.Now);
            foreach (var record in records)
            {
                var cell_props = this.get_rw_string_props(record.type);
                fp.WriteLine();
                fp.WriteLine();
                fp.WriteLine("// " + record.type.Name);
                fp.WriteLine();
                fp.WriteLine("public class {0}", record.type.Name);
                fp.WriteLine("{");
                foreach (var cellprop in cell_props)
                {
                    fp.WriteLine("    public string {0} {{get; set;}}", cellprop.Name);                    
                }
                fp.WriteLine("}");
                fp.WriteLine();
            }

            fp.WriteLine("public class MetadataDB");
            fp.WriteLine("{");
            foreach (var record in records)
            {
                string itemname = record.type.Name.ToLower();
                string colname = record.type.Name + "s";
                fp.WriteLine("   public List<{0}> {1} {{get;set;}}", colname, record.type.Name);
            }

            fp.WriteLine("public MetadataDB()");
            fp.WriteLine("{");
            foreach (var record in records)
            {
                var cell_props = this.get_rw_string_props(record.type);
                string itemname = record.type.Name.ToLower();
                string colname = record.type.Name + "s";
                fp.WriteLine("    this.{0} = new List<{1}>();", colname, record.type.Name);

                foreach (var cell in record.col)
                {
                    fp.Write("    {0}.Add(", colname);
                    var x = get_ctor(record.type, cell, cell_props);
                    fp.Write(x);
                    fp.WriteLine(");");
                    fp.Flush();
                }
            }
            fp.WriteLine("}");
            fp.WriteLine("}");

            fp.Close();
        }

        private static string get_ctor(Type itemType, object cell, List<PropertyInfo> cell_props)
        {
            var names = cell_props.Select(p => p.Name).ToList();
            var values = cell_props.Select(p => p.GetValue(cell, null) ?? "").ToList();
            var pairs = Enumerable.Range(0, names.Count).Select(i => names[i] + "=\"" + values[i] + "\"").ToArray();
            string ctor_string = string.Format("new {0} {{ {1} }}", itemType.Name, string.Join(", ", pairs));
            return ctor_string;
        }

        public List<System.Reflection.PropertyInfo> get_rw_string_props(System.Type item_type)
        {
            var bf = System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance;
            ;
            var properties = item_type.GetProperties(bf);
            var target_props = properties.Where(p => p.CanRead).Where(p => p.PropertyType == typeof (string)).ToList();
            return target_props;
        }
    }
}