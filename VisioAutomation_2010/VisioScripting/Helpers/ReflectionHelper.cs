using System.Linq;

namespace VisioScripting.Helpers
{
    internal class ReflectionHelper
    {
        public static string GetTypeCategoryDisplayString(TypeCategory type)
        {
            if (type == TypeCategory.StaticClass)
            {
                return "static class";
            }
            else if (type == TypeCategory.AbstractClass)
            {
                return "abstract class";
            }
            else if (type == TypeCategory.Class)
            {
                return "class";
            }
            else if (type == TypeCategory.Enum)
            {
                return "enum";
            }
            else if (type == TypeCategory.Interface)
            {
                return "interface";
            }
            else if (type == TypeCategory.Struct)
            {
                return "struct";
            }
            else
            {
                return string.Empty;
            }
        }
        
        public static string GetTypeCategoryDisplayString(System.Type type)
        {
            var cat = ReflectionHelper.GetTypeCategory(type);
            return ReflectionHelper.GetTypeCategoryDisplayString(cat);
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
                if (ReflectionHelper.TypeIsStaticClass(type))
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
            else if (ReflectionHelper.TypeIsStruct(type))
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
            options.NameOverrideFunc = ReflectionHelper.GetCSharpTypeAlias;
            return ReflectionHelper.GetNiceTypeName(type, options);
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

            if (ReflectionHelper.IsNullableType(type))
            {
                var actualtype = type.GetGenericArguments()[0];
                return ReflectionHelper.GetNiceTypeName(actualtype, options) + "?";
            }

            if (type.IsArray)
            {
                var at = type.GetElementType();
                return string.Format("{0}[]", ReflectionHelper.GetNiceTypeName(at, options));
            }

            if (type.IsGenericType)
            {
                var sb = new System.Text.StringBuilder();
                var tokens = type.Name.Split('`');

                sb.Append(tokens[0]);
                var gas = type.GetGenericArguments();
                var ga_names = gas.Select(i => ReflectionHelper.GetNiceTypeName(i, options));

                sb.Append("<");
                sb.Append( string.Join(", ",ga_names) );
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
    }
}