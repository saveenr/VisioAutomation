using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.Scripting.Commands
{
    internal class ReflectionUtil
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
                return "";
            }
        }
        
        public static string GetTypeCategoryDisplayString(System.Type type)
        {
            var cat = ReflectionUtil.GetTypeCategory(type);
            return ReflectionUtil.GetTypeCategoryDisplayString(cat);
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
                if (ReflectionUtil.TypeIsStaticClass(type))
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
            else if (ReflectionUtil.TypeIsStruct(type))
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
            options.NameOverrideFunc = ReflectionUtil.GetCSharpTypeAlias;
            return ReflectionUtil.GetNiceTypeName(type, options);
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

            if (ReflectionUtil.IsNullableType(type))
            {
                var actualtype = type.GetGenericArguments()[0];
                return ReflectionUtil.GetNiceTypeName(actualtype, options) + "?";
            }

            if (type.IsArray)
            {
                var at = type.GetElementType();
                return $"{ReflectionUtil.GetNiceTypeName(at, options)}[]";
            }

            if (type.IsGenericType)
            {
                var sb = new System.Text.StringBuilder();
                var tokens = type.Name.Split('`');

                sb.Append(tokens[0]);
                var gas = type.GetGenericArguments();
                var ga_names = gas.Select(i => ReflectionUtil.GetNiceTypeName(i, options));

                sb.Append("<");
                sb.AppendJoin(", ", ga_names);
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

    internal static class StringBuilderExtensions
    {
        public static void AppendJoin(this System.Text.StringBuilder sb, string s, IEnumerable<string> tokens)
        {
            // This works exactly like string.Join - except that it appends the results of the join into
            // a StringBuilder object

            // First, make sure the stringbuilder has enough capacity allocated 
            int num_tokens = 0;
            int tokens_length  = 0;

            // use a foreach to minimize the times we have to go through the enumerable
            foreach (string token in tokens)
            {
                num_tokens++;
                tokens_length += token.Length;
            }

            // figure out how much space is needed for the separators
            int num_seps = (num_tokens > 1) ? num_tokens - 1 : 0;
            int separators_length = s.Length*num_seps;

            int combined_length = tokens_length + separators_length;
            int required_capacity = combined_length + sb.Length;
            sb.EnsureCapacity(required_capacity);

            // Now add the tokens and separators
            int i = 0;
            foreach (string token in tokens)
            {
                if (i > 0)
                {
                    sb.Append(s);
                }
                sb.Append(token);
                i++;
            }
        }
    }
}