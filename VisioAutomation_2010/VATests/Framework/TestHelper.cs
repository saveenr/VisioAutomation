using System;
using System.IO;

namespace VisioAutomation_Tests
{
    public class TestHelper
    {
        private readonly string _output_path;

        public TestHelper(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new ArgumentException("name is null or empty", nameof(name));
            }

            string mydocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            this._output_path = Path.Combine(mydocs, name);

            if (!Directory.Exists(this._output_path))
            {
                Directory.CreateDirectory(this._output_path);
            }
        }

        public string GetOutputFilename(string method, string ext)
        {
            if (ext == null)
            {
                throw new System.ArgumentNullException(nameof(ext));
            }

            if (ext.Length < 1)
            {
                throw new System.ArgumentException(nameof(ext));
            }

            if (ext[0] != '.')
            {
                throw new System.ArgumentException(nameof(ext));
            }

            string abs_path = this._output_path;
            var culture = System.Globalization.CultureInfo.InvariantCulture;
            var datetime_str = DateTime.Now.ToString("yyyy-dd-M--HH-mm-ss", culture);
            var basename = method + "_" + datetime_str + ext;
            string abs_filename = Path.Combine(abs_path, basename);
            return abs_filename;
        }
    }
}
