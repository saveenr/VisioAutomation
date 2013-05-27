using System.Collections.Generic;
using System.Linq;

namespace TestCommon
{
    public class Helper
    {
        private readonly string OutputPath;

        public Helper(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new System.ArgumentException("name");
            }

            this.OutputPath = GetOutputPathEx(name);

            PrepareOutputPath();
        }

        public string GetTestMethodOutputFilename(string ext)
        {
            string abs_path = this.OutputPath;
            string abs_filename = System.IO.Path.Combine(abs_path, GetMethodName(2) + ext);
            return abs_filename;
        }

        public string GetTestMethodOutputFilename()
        {
            string abs_path = this.OutputPath;
            string abs_filename = System.IO.Path.Combine(abs_path, GetMethodName(2));
            return abs_filename;
        }

        private void PrepareOutputPath()
        {
            if (!System.IO.Directory.Exists(this.OutputPath))
            {
                System.IO.Directory.CreateDirectory(this.OutputPath);
            }
        }

        private static string GetOutputPathEx(string name)
        {
            string path = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments);
            return System.IO.Path.Combine(path, name);
        }

        private static string GetMethodName(int depth)
        {
            var stackTrace = new System.Diagnostics.StackTrace();
            var stackFrame = stackTrace.GetFrame(depth);
            var methodBase = stackFrame.GetMethod();

            string n = methodBase.DeclaringType.Name + "." + methodBase.Name;
            return n;
        }

        private static string GetMethodName()
        {
            return GetMethodName(2);
        }

        private static string GetMethodName(string ext)
        {
            return GetMethodName(2) + ext;
        }


        public static Dictionary<string, T> EnumToDictionary<T>(System.Type t)
        {
            var dic = new Dictionary<string, T>();
            string[] names = System.Enum.GetNames(t);
            System.Array avalues = System.Enum.GetValues(t);
            for (int i = 0; i < avalues.Length; i++)
            {

                dic[names[i]] = (T)avalues.GetValue(i);
            }

            return dic;
        }

        public static List<T> GetDuplicates<T>(IEnumerable<T> items)
        {
            var set = new HashSet<T>();
            var dupes = new List<T>();

            foreach (var item in items)
            {
                if (set.Contains(item))
                {
                    dupes.Add(item);
                }
                else
                {
                    set.Add(item);
                }
            }

            return dupes;
        }

        public static void AssertNoDuplicates<T>(IEnumerable<T> items)
        {
            var set = new HashSet<T>();
            var dupes = new List<T>();

            foreach (var item in items)
            {
                if (set.Contains(item))
                {
                    dupes.Add(item);
                }
                else
                {
                    set.Add(item);
                }
            }

            if (dupes.Count > 0)
            {
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.Fail("Duplicated {0}", dupes.Count);
            }
        }

        public const string LoremIpsumText =
             @"Lorem ipsum dolor sit amet, consectetur adipisicing elit,
sed do eiusmod tempor incididunt ut labore et dolore magna
aliqua. Ut enim ad minim veniam, quis nostrud exercitation 
ullamco laboris nisi ut aliquip ex ea commodo consequat. 
Duis aute irure dolor in reprehenderit in voluptate velit 
esse cillum dolore eu fugiat nulla pariatur. Excepteur sint
occaecat cupidatat non proident, sunt in culpa qui officia
deserunt mollit anim id est laborum";
    }
}