using System;
using System.Diagnostics;
using System.IO;

namespace TestVisioAutomationVDX
{
    public class TestHelper
    {
        private readonly string OutputPath;

        public TestHelper(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new ArgumentException("name is null or empty","name");
            }

            this.OutputPath = GetOutputPathEx(name);

            this.PrepareOutputPath();
        }

        public string GetTestMethodOutputFilename(string ext)
        {
            string abs_path = this.OutputPath;
            string abs_filename = Path.Combine(abs_path, GetMethodName(2) + DateTime.Now.ToString("yyyy-dd-M--HH-mm-ss") + ext);
            return abs_filename;
        }

        public string GetTestMethodOutputFilename()
        {
            string abs_path = this.OutputPath;
            string abs_filename = Path.Combine(abs_path, GetMethodName(2));
            return abs_filename;
        }

        private void PrepareOutputPath()
        {
            if (!Directory.Exists(this.OutputPath))
            {
                Directory.CreateDirectory(this.OutputPath);
            }
        }

        private static string GetOutputPathEx(string name)
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            return Path.Combine(path, name);
        }

        private static string GetMethodName(int depth)
        {
            var stackTrace = new StackTrace();
            var stackFrame = stackTrace.GetFrame(depth);
            var methodBase = stackFrame.GetMethod();

            string n = methodBase.DeclaringType.Name + "." + methodBase.Name;
            return n;
        }

    }
}