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

            string dirname = FilterDirectory(files[0]);

            string docpath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string myTempFile = GenerateFileName(docpath, dirname, ".csv");
            using (StreamWriter sw = new StreamWriter(myTempFile))
            {
                string prefix = "";
                for (int i = 0; i < Properties.Settings.Default.skip_cols; i++)
                    prefix += ",";


                for (int i = 0; i < Properties.Settings.Default.skip_rows-1; i++)
                    sw.WriteLine("");

                sw.WriteLine(string.Format("{0}\"{1}\"", prefix, dirname));

                sw.WriteLine("");

                foreach (string file in files)
                    sw.WriteLine(string.Format("{0}\"{1}\"", prefix, FilterFilename(file)));
            }
            string cmdargs = string.Format("/c start {0}", myTempFile);
            Process.Start("cmd.exe", cmdargs);
            WindowState = FormWindowState.Minimized;
        }

        private string GenerateFileName(string docpath, string context, string extension)
        {
            string pattern = "yyyy-MM-dd_HHmmss";
            string ret = Path.Combine(docpath, string.Format("{0}_{1}.csv", context, DateTime.Now.ToString(pattern)));
            int counter = 1;
            while (File.Exists(ret))
            {
                counter += 1;
                ret = Path.Combine(docpath, string.Format("{0}_{1}({2}).csv", context, DateTime.Now.ToString(pattern), counter));
            }
            ret = ret.Replace(' ', '_');
            return ret;
        }

        private String FilterDirectory(string infile)
        {
            string ret = infile;
            if (infile.Contains("\\"))
                ret = infile.Substring(0, infile.LastIndexOf("\\"));
            if (ret.Contains("\\"))
                ret = ret.Substring(ret.LastIndexOf("\\")+1);
            else
                ret = "";
            return ret;
        }

        private string FilterFilename(string infile)
        {
            string ret = infile;
            //if (infile.Contains("Dropbox"))
            //    ret = infile.Substring(infile.IndexOf("Dropbox"));
            if (infile.Contains("\\"))
                ret = infile.Substring(infile.LastIndexOf("\\") + 1);

            return ret;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.IsMaximized = WindowState == FormWindowState.Maximized;
            Properties.Settings.Default.WindowPosition = DesktopBounds;
            Properties.Settings.Default.Save();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Drag-and-drop filenames from Windows Explorer to create an Excel Spreadsheet\n\nMIT Licensed\nFor more information see <https://github.com/rstms/ExcelFileList>", "About ExcelFileList");
        }

        private DialogResult ShowInputDialog(ref string input, string title)
        {
            System.Drawing.Size size = new System.Drawing.Size(200, 70);
            Form inputBox = new Form();

            inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            inputBox.MinimizeBox = false;
            inputBox.MaximizeBox = false;
            inputBox.ClientSize = size;
            inputBox.Text = title;

            System.Windows.Forms.TextBox textBox = new TextBox();
            textBox.Size = new System.Drawing.Size(size.Width - 10, 23);
            textBox.Location = new System.Drawing.Point(5, 5);
            textBox.Text = input;
            inputBox.Controls.Add(textBox);

            Button okButton = new Button();
            okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            okButton.Name = "okButton";
            okButton.Size = new System.Drawing.Size(75, 23);
            okButton.Text = "&OK";
            okButton.Location = new System.Drawing.Point(size.Width - 80 - 80, 39);
            inputBox.Controls.Add(okButton);

            Button cancelButton = new Button();
            cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            cancelButton.Name = "cancelButton";
            cancelButton.Size = new System.Drawing.Size(75, 23);
            cancelButton.Text = "&Cancel";
            cancelButton.Location = new System.Drawing.Point(size.Width - 80, 39);
            inputBox.Controls.Add(cancelButton);

            inputBox.AcceptButton = okButton;
            inputBox.CancelButton = cancelButton;

            inputBox.StartPosition = FormStartPosition.Manual;
            inputBox.Location = new Point(this.Location.X + this.ClientSize.Width / 2 - inputBox.ClientSize.Width / 2, this.Location.Y);

            DialogResult result = inputBox.ShowDialog(this);
            input = textBox.Text;
            return result;
        }

        private void skiprowsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string input = Properties.Settings.Default.skip_rows.ToString();

            DialogResult result = ShowInputDialog(ref input, "Number of rows to skip");
            if (result.Equals(DialogResult.OK))
            {
                Properties.Settings.Default.skip_rows = Int32.Parse(input);
                Properties.Settings.Default.Save();
            }
        }

        private void skipcolsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string input = Properties.Settings.Default.skip_cols.ToString();
            DialogResult result = ShowInputDialog(ref input, "Number of columns to skip");
            if (result.Equals(DialogResult.OK))
            {
                Properties.Settings.Default.skip_cols = Int32.Parse(input);
                Properties.Settings.Default.Save();
            }

        }
    }
}

