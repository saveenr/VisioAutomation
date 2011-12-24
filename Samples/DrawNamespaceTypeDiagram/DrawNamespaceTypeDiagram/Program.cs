using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace DrawNamespaceTypeDiagram
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("Syntax: DrawNamespaceTypeDiagram.exe <filename.dll>");
                System.Environment.Exit(-1);
            }

            string filename = System.IO.Path.GetFullPath(args[0]);

            var asm = System.Reflection.Assembly.LoadFile(filename);
            var app = new IVisio.ApplicationClass();
            var om_containers = CreateContainerModelFromAssembly(asm);
            om_containers.LayoutOptions.RenderWithShapes = true;
            om_containers.Render(app);
        }

        private static VA.Layout.ContainerLayout.ContainerModel CreateContainerModelFromAssembly(System.Reflection.Assembly asm)
        {
            var types = asm.GetExportedTypes().Where(t => t.IsPublic);

            var ns_dic = new Dictionary<string, List<System.Type>>();
            foreach (var t in types)
            {
                List<System.Type> list;
                string ns = t.Namespace;
                if (ns_dic.ContainsKey(ns))
                {
                    list = ns_dic[ns];
                }
                else
                {
                    list = new List<Type>();
                    ns_dic[ns] = list;
                }
                list.Add(t);
            }

            bool show_type_kind = true;

            var om_containers = new VA.Layout.ContainerLayout.ContainerModel();

            var sorted_namespaces = ns_dic.Keys.OrderBy(i => i).ToList();

            foreach (string ns in sorted_namespaces)
            {
                var nstypes = ns_dic[ns];

                var sorted_items =
                    nstypes.Select(t => new
                                            {
                                                type = t, 
                                                kind = get_type_kindEx(t), 
                                                name = get_nice_type_name(t)
                                            }).OrderBy(i => i.kind).
                        ThenBy(i => i.name).ToList();

                Console.WriteLine("{0}", ns);

                var om_container = om_containers.AddContainer(ns);

                om_container.FillForegnd = "rgb(240,240,240)";
                om_container.LineWeight = "0";
                om_container.LinePattern = "0";
                om_container.VerticalAlign = "0";

                foreach (var i in sorted_items)
                {
                    Console.WriteLine("    {0}", i.name);
                    var item = om_container.Add( get_type_kindname(i.kind) + " " + i.name);

                    item.FillForegnd = "rgb(255,255,255)";
                    item.LineWeight = "0";
                    item.LinePattern = "0";

                    if (i.kind == TypeKind.Enum)
                    {
                        item.FillForegnd = "rgb(220,240,245)";
                    }
                }
            }
            return om_containers;
        }

        private static string get_nice_type_name(System.Type type)
        {
            if (type.IsGenericType)
            {
                var sb = new System.Text.StringBuilder();
                var tokens = type.Name.Split(new[] {'`'});


                sb.Append(tokens[0]);
                var gas = type.GetGenericArguments();
                var ga_names = gas.Select(i => i.Name).ToList();

                sb.Append("<");
                sb.Append(string.Join(", ", ga_names));
                sb.Append(">");
                return sb.ToString();
            }

            return type.Name;
        }

        private static string get_type_kindname(TypeKind type)
        {
            if (type == TypeKind.StaticClass)
            {
                return "static class";
            }
            else if (type ==TypeKind.AbstractClass)
            {
                return "abstract class";
            }
            else if( type ==TypeKind.Class)
            {
                return "class";
            }
            else if (type == TypeKind.Enum)
            {
                return "enum";
            }
            else if (type == TypeKind.Interface)
            {
                return "interface";
            }
            else if (type == TypeKind.Struct)
            {
                return "struct";
            }
            else
            {
                return "";
            }
        }

        private static TypeKind get_type_kindEx(System.Type type)
        {
            if (type.IsClass)
            {
                if (TypeIsStaticClass(type))
                {
                    return TypeKind.StaticClass;
                }
                else if (type.IsAbstract)
                {
                    return TypeKind.AbstractClass;
                }
                return TypeKind.Class;
            }
            else if (type.IsEnum)
            {
                return TypeKind.Enum;
            }
            else if (type.IsInterface)
            {
                return TypeKind.Interface;
            }
            else if (TypeIsStruct(type))
            {
                return TypeKind.Struct;
            }
            else
            {
                return TypeKind.Other;
            }
        }


        private static bool TypeIsStruct(System.Type type)
        {
            return (type.IsValueType && !type.IsPrimitive && !type.Namespace.StartsWith("System") && !type.IsEnum);
        }

        private static bool TypeIsStaticClass(System.Type type)
        {
            return (type.IsAbstract && type.IsSealed);
        }
    }

    public enum TypeKind
    {
        StaticClass,
        Class,
        AbstractClass,
        Interface,
        Struct,
        Enum,
        Other
    }
}

