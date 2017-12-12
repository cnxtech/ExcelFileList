using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelFileList
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            AllowDrop = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.IsMaximized)
                WindowState = FormWindowState.Maximized;
            else if (Screen.AllScreens.Any(screen => screen.WorkingArea.IntersectsWith(Properties.Settings.Default.WindowPosition)))
            {
                StartPosition = FormStartPosition.Manual;
                DesktopBounds = Properties.Settings.Default.WindowPosition;
                WindowState = FormWindowState.Minimized;
            }
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            string docpath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string myTempFile = Path.Combine(docpath, GenerateFileName("files", ".csv"));
            using (StreamWriter sw = new StreamWriter(myTempFile))
            {
                sw.WriteLine("");
                sw.WriteLine("");
                foreach (string file in files)
                    sw.WriteLine(string.Format(",{0}", FilterFilename(file)));
            }
            string cmdargs = string.Format("/c start {0}", myTempFile);
            Process.Start("cmd.exe", cmdargs);
            WindowState = FormWindowState.Minimized;
        }

        private string GenerateFileName(string context, string extension)
        {
            return context + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + "_" + Guid.NewGuid().ToString("N") + extension;

        }

        private string FilterFilename(string infile)
        {
            string ret = infile;
            if (infile.Contains("Dropbox"))
                ret = infile.Substring(infile.IndexOf("Dropbox"));
            return ret;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.IsMaximized = WindowState == FormWindowState.Maximized;
            Properties.Settings.Default.WindowPosition = DesktopBounds;
            Properties.Settings.Default.Save();
        }
    }
}
