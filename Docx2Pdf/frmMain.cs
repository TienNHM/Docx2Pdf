using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace Docx2Pdf
{
    public partial class frmMain : Form
    {
        public Document Document { get; set; }
        private OpenFileDialog files;
        public frmMain()
        {
            InitializeComponent();
            btStart.Enabled = false;
            btEnd.Enabled = false;
        }

        private void btSelect_Click(object sender, EventArgs e)
        {
            progressBar.Value = 0;
            lbStatus.Text = "0 %";
            files = new OpenFileDialog()
            {
                Multiselect = true
            };
            if (files.ShowDialog() == DialogResult.OK)
                btStart.Enabled = true;
        }

        private void btStart_Click(object sender, EventArgs e)
        {
            btStart.Enabled = false;
            btEnd.Enabled = true;
            DialogResult start = MessageBox.Show("Start this process?", "Info", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (start == DialogResult.Yes)
                backgroundWorker.RunWorkerAsync();
        }
        private void btEnd_Click(object sender, EventArgs e)
        {
            DialogResult dialog = MessageBox.Show("Do you want to end this task?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dialog == DialogResult.Yes)
            {
                lbStatus.Text = "Please wait for killing this task...";
                backgroundWorker.CancelAsync();
            }
        }

        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            app.ShowWindowsInTaskbar = false;
            string[] listFiles = files.FileNames;
            int i = 0;
            foreach (var path in listFiles)
            {
                backgroundWorker.ReportProgress((++i * 100) / listFiles.Length);
                Document = app.Documents.Open(FileName: @path);
                string name = path.Substring(0, path.LastIndexOf('.'));
                Document.ExportAsFixedFormat(@name + ".pdf", WdExportFormat.wdExportFormatPDF);
                app.Documents.Close();
                if (backgroundWorker.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }
            }
            app.Quit();
            backgroundWorker.ReportProgress(100);
        }

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
            lbStatus.Text = progressBar.Value + " %";
        }
        
        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btEnd.Enabled = false;
            if (e.Cancelled)
            {
                lbStatus.Text = progressBar.Value + " %";
                MessageBox.Show("Task cancelled!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (e.Error != null)
                MessageBox.Show("Error(s) occurred!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                MessageBox.Show("Done!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                progressBar.Value = 0;
                lbStatus.Text = "0 %";
            }
        }
    }
}
