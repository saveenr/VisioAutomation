using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.Scripting.Commands
{
    internal class ReflectionUtil
    {
        public static string GetTypeCategoryDisplayString(ReflectionUtil.TypeCategory type)
        {
            if (type == ReflectionUtil.TypeCategory.StaticClass)
            {
                return "static class";
            }
            else if (type == ReflectionUtil.TypeCategory.AbstractClass)
            {
                return "abstract class";
            }
            else if (type == ReflectionUtil.TypeCategory.Class)
            {
                return "class";
            }
            else if (type == ReflectionUtil.TypeCategory.Enum)
            {
                return "enum";
            }
            else if (type == ReflectionUtil.TypeCategory.Interface)
            {
                return "interface";
            }
            else if (type == ReflectionUtil.TypeCategory.Struct)
            {
                return "struct";
            }
            else
            {
                return "";
            }
        }
        
        public static string GetTypeCategoryDisplayString(System.Type type)
        {
            var cat = GetTypeCategory(type);
            return GetTypeCategoryDisplayString(cat);
        }

        private static bool TypeIsStruct(System.Type type)
        {
            return (type.IsValueType && !type.IsPrimitive && !type.Namespace.StartsWith("System") && !type.IsEnum);
        }

        private static bool TypeIsStaticClass(System.Type type)
        {
            return (type.IsAbstract && type.IsSealed);
        }
        
        public enum TypeCategory
        {
            StaticClass,
            Class,
            AbstractClass,
            Interface,
            Struct,
            Enum,
            Other
        }
 
        public static TypeCategory GetTypeCategory(System.Type type)
        {
            if (type.IsClass)
            {
                if (TypeIsStaticClass(type))
                {
                    return TypeCategory.StaticClass;
                }
                else if (type.IsAbstract)
                {
                    return TypeCategory.AbstractClass;
                }
                return TypeCategory.Class;
            }
            else if (type.IsEnum)
            {
                return TypeCategory.Enum;
            }
            else if (type.IsInterface)
            {
                return TypeCategory.Interface;
            }
            else if (TypeIsStruct(type))
            {
                return TypeCategory.Struct;
            }
            else
            {
                return TypeCategory.Other;
            }
        }


        public static string GetCSharpTypeAlias(System.Type type)
        {
            if (type == typeof(int))
            {
                return "int";
            }
            else if (type == typeof(string))
            {
                return "string";
            }
            else if (type == typeof(double))
            {
                return "double";
            }
            else if (type == typeof(bool))
            {
                return "bool";
            }
            else if (type == typeof(short))
            {
                return "short";
            }
            else if (type == typeof(ushort))
            {
                return "ushort";
            }
            else if (type == typeof(decimal))
            {
                return "decimal";
            }
            else if (type == typeof(double))
            {
                return "double";
            }
            else if (type == typeof(float))
            {
                return "float";
            }
            else if (type == typeof(char))
            {
                return "char";
            }
            else if (type == typeof(uint))
            {
                return "uint";
            }
            else if (type == typeof(long))
            {
                return "long";
            }
            else if (type == typeof(ulong))
            {
                return "ulong";
            }
            else if (type == typeof(byte))
            {
                return "byte";
            }
            else if (type == typeof(sbyte))
            {
                return "sbyte";
            }
            else if (type == typeof(object))
            {
                return "object";
            }
            else
            {
                return null;
            }
        }

        public class NamingOptions
        {
            public System.Func<System.Type, string> NameOverrideFunc;
        }

        public static string GetNiceTypeName(System.Type type)
        {
            var options = new NamingOptions();
            options.NameOverrideFunc = GetCSharpTypeAlias;
            return GetNiceTypeName(type, options);
        }

        public static string GetNiceTypeName(System.Type type, NamingOptions options)
        {
            if (options != null && options.NameOverrideFunc !=null)
            {
                string s = options.NameOverrideFunc(type);
                if (s != null)
                {
                    return s;
                }
            }

            if (IsNullableType(type))
            {
                var actualtype = type.GetGenericArguments()[0];
                return GetNiceTypeName(actualtype, options) + "?";
            }

            if (type.IsArray)
            {
                var at = type.GetElementType();
                return string.Format("{0}[]", GetNiceTypeName(at, options));
            }

            if (type.IsGenericType)
            {
                var sb = new System.Text.StringBuilder();
                var tokens = type.Name.Split(new[] { '`' });

                sb.Append(tokens[0]);
                var gas = type.GetGenericArguments();
                var ga_names = gas.Select(i => GetNiceTypeName(i, options));

                sb.Append("<");
                Join(sb, ", ", ga_names);
                sb.Append(">");
                return sb.ToString();
            }

            return type.Name;
        }

        public static bool IsNullableType(System.Type colType)
        {
            return ((colType.IsGenericType) &&
                    (colType.GetGenericTypeDefinition() == typeof(System.Nullable<>)));
        }

        private static void Join(System.Text.StringBuilder sb, string s, IEnumerable<string> tokens)
        {
            int n = tokens.Count();
            int c = tokens.Select(t => t.Length).Sum();
            c += (n > 1) ? s.Length * n : 0;
            c += sb.Length;
            sb.EnsureCapacity(c);

            int i = 0;
            foreach (string t in tokens)
            {
                if (i > 0)
                {
                    sb.Append(s);
                }
                sb.Append(t);
                i++;
            }
        }   

    }
}