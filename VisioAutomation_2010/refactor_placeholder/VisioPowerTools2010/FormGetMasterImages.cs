using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Visio;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerTools2010
{
    public partial class FormGetMasterImages : Form
    {
        private string output_basename = "Catalog.htm";

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        extern static bool DestroyIcon(IntPtr handle);

        [DllImport("gdi32.dll")]
        static extern bool DeleteEnhMetaFile(IntPtr hemf);

        public FormGetMasterImages()
        {
            this.InitializeComponent();
        }

        private void log(string fmt, params object[] tokens)
        {
            var s = string.Format(fmt, tokens);
            var snl = string.Format("{0}\n", s);
            this.textBoxLog.AppendText(snl);
        }

        private string get_src_folder()
        {
            return this.normalize_path(this.textBoxStencilFolder.Text);
        }

        private string get_dest_folder()
        {
            return this.normalize_path( this.textBoxOutputFolder.Text );
        }

        private string normalize_path(string s)
        {
            while (s.EndsWith(@"\"))
            {
                s = s.Substring(0, s.Length - 1);
            }
            return s;
        }

        private void buttonRun_Click(object sender, EventArgs e)
        {
            this.textBoxLog.Clear();
            string srcfolder = this.get_src_folder();

            if (!System.IO.Directory.Exists(srcfolder))
            {
                MessageBox.Show("Source folder does not exist");
                return;
            }

            if (!System.IO.Path.IsPathRooted(srcfolder))
            {
                MessageBox.Show("Source folder is not an absolute path");
                return;
            }

            string destfolder = this.get_dest_folder();

            if (!System.IO.Path.IsPathRooted(destfolder))
            {
                MessageBox.Show("Output folder is not an absolute path");
                return;
            }

            var app = Globals.ThisAddIn.Application;

            var ver = VisioAutomation.Application.ApplicationHelper.GetVersion(app);

            this.log("Searching for Stencil files (VSS, VSSX)");
            var stencilfiles = System.IO.Directory.GetFiles(srcfolder, "*.vss").ToList();

            // If Visio 2013 then check for VSSX files
            if (ver.Major == 15)
            {
                var vssx_stencilfiles = System.IO.Directory.GetFiles(srcfolder, "*.vssx").ToList();

                stencilfiles.AddRange(vssx_stencilfiles);
            }

            this.log("Found {0} stencil files", stencilfiles.Count);

            this.log("Starting Visio Application");
            var docs = app.Documents;

            try
            {
                this.create_folder_safe(destfolder);
            }
            catch (Exception)
            {
                return;
            }

            if (!System.IO.Directory.Exists(destfolder))
            {
                return;
            }

            string html_filename = System.IO.Path.Combine(this.get_dest_folder(), this.output_basename);

            try
            {
                using (var writer = new SimpleHtml5Writer(html_filename))
                {
                    writer.DocType("HTML5");
                    writer.Start("html");

                    writer.Start("head");
                    writer.Start("style");
                    writer.Text(".stencilname { font-family: \"Segoe UI Light\"; font-size:30pt}");
                    writer.Text(".mastername { font-family: \"Segoe UI\"; font-size:10pt}");
                    writer.Text("td { padding-bottom: 50pt;");
                    writer.End("style");
                    writer.End("head");

                    writer.Start("body");


                    foreach (var stencilfilename in stencilfiles)
                    {
                        writer.Start("table");
                        var stencilfilename_basename = System.IO.Path.GetFileName(stencilfilename);

                        writer.Start("tr");
                        writer.Start("td");
                        writer.Attribute("colspan", "3");
                        writer.Attribute("class", "stencilname");
                        writer.Text(stencilfilename_basename);
                        writer.End("td");
                        writer.End("tr");

                        var stencilfilename_basename_wo_ext = System.IO.Path.GetFileNameWithoutExtension(stencilfilename);
                        this.log("Loading \"{0}\"", stencilfilename_basename);
                        var doc = docs.Add(stencilfilename);

                        string stencilname_safe = FormGetMasterImages.MakeSafeFilename(stencilfilename_basename_wo_ext, '_');
                        string cur_destfolder = System.IO.Path.Combine(destfolder, stencilname_safe);

                        try
                        {
                            this.create_folder_safe(cur_destfolder);
                        }
                        catch (Exception)
                        {
                            return;
                        }

                        var masters = doc.Masters;
                        int num_masters = masters.Count;

                        for (int i = 1; i <= num_masters; i++)
                        {
                            writer.Start("tr");
                            writer.Attribute("style", "vertical-align:top");
                            var master = masters[i];
                            this.log("    master {0}", master.Name);
                            string mastername_safe = FormGetMasterImages.MakeSafeFilename(master.Name, '_');

                            string picture_dir = System.IO.Path.Combine(cur_destfolder, "pictures");
                            string icon_dir = System.IO.Path.Combine(cur_destfolder, "icons");

                            this.create_folder_safe(picture_dir);
                            this.create_folder_safe(icon_dir);

                            string picture_filename = System.IO.Path.Combine(picture_dir, mastername_safe + ".png");
                            string icon_filename = System.IO.Path.Combine(icon_dir, mastername_safe + ".png");

                            if (!System.IO.File.Exists(icon_filename))
                            {
                                FormGetMasterImages.SaveMasterIcon(icon_filename, master);
                            }
                            else
                            {
                                this.log("        icon PNG already exists. Skipping.");
                            }

                            if (!System.IO.File.Exists(picture_filename))
                            {
                                FormGetMasterImages.SaveMasterPicture(picture_filename, master);
                            }
                            else
                            {
                                this.log("        picture PNG already exists. Skipping.");
                            }

                            writer.Start("td");
                            writer.Attribute("width", "200");
                            writer.Attribute("class", "mastername");
                            writer.Text(master.NameU);
                            writer.End("td");

                            writer.Start("td");
                            writer.Attribute("width", "150");
                            writer.Start("img");
                            string icon_src = icon_filename.Substring(destfolder.Length + 1);
                            writer.Attribute("src", icon_src);
                            writer.End("img");
                            writer.End("td");

                            writer.Start("td");
                            writer.Attribute("width", "250");
                            writer.Start("img");
                            string picture_src = picture_filename.Substring(destfolder.Length + 1);
                            writer.Attribute("src", picture_src);
                            writer.End("img");
                            writer.End("td");

                            writer.End("tr");
                        }

                        this.log("Closing stencil doc");
                        doc.Close();
                        writer.End("table");

                    }

                    this.log("Finished.");

                    writer.End("body");
                    writer.End("html");                    
                }

            }
            catch (Exception)
            {
                this.log("Could not create file \"{0}\"", html_filename);
                return;
            }

        }

        private static void SaveMasterPicture(string picture_filename, Master master)
        {
            stdole.IPicture master_picture_pic = (stdole.IPicture) master.Picture;
            IntPtr metafile_handle = (IntPtr) master_picture_pic.Handle;
            using (var metafile = new System.Drawing.Imaging.Metafile(metafile_handle, true))
            {
                metafile.Save(picture_filename);
            }
            FormGetMasterImages.DeleteEnhMetaFile(metafile_handle);
        }

        private static void SaveMasterIcon(string icon_filename, Master master)
        {
            stdole.IPicture master_icon_pic = (stdole.IPicture) master.Icon;
            IntPtr icon_handle = (IntPtr) master_icon_pic.Handle;
            using (var icon = System.Drawing.Icon.FromHandle(icon_handle).ToBitmap())
            {
                icon.Save(icon_filename);
            }
            FormGetMasterImages.DestroyIcon(icon_handle);
        }

        private void create_folder_safe(string cur_destfolder)
        {
            if (!System.IO.Directory.Exists(cur_destfolder))
            {
                this.log("Creating output folder");
                try
                {
                    System.IO.Directory.CreateDirectory(cur_destfolder);
                }
                catch (System.IO.IOException)
                {
                    this.log("Failed to create directory \"{0}\"", cur_destfolder);
                    throw;
                }
            }
        }

        public static string MakeSafeFilename(string filename, char replaceChar)
        {
            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
            {
                filename = filename.Replace(c, replaceChar);
            }
            return filename;
        }

        private void FormGetMasterImages_Load(object sender, EventArgs e)
        {
            this.textBoxStencilFolder.Text = Properties.Settings.Default.StencilCatalogInputFolder;
            this.textBoxOutputFolder.Text  = Properties.Settings.Default.StencilCatalogOutputFolder;

        }

        private void linkLabelOpenOutput_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string html_filename = System.IO.Path.Combine(this.get_dest_folder(), this.output_basename);

            if (System.IO.File.Exists(html_filename))
            {
                System.Diagnostics.Process.Start(html_filename);
                
            }
            else
            {
                MessageBox.Show("Output html file does not exist");
            }
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Close();

            this.SaveSettings();
        }

        private void SaveSettings()
        {
            Properties.Settings.Default.StencilCatalogInputFolder = this.textBoxStencilFolder.Text;
            Properties.Settings.Default.StencilCatalogOutputFolder = this.textBoxOutputFolder.Text;
        }

        private void FormGetMasterImages_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.SaveSettings();
        }
    }
}
