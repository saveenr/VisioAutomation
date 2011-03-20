using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Linq;

namespace TestVisioAutomation
{
    public class TestHelper
    {
        public readonly string OutputPath;
        public TestHelper(string name)
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

    }
}

