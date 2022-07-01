using System;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;

namespace ExcelImageDownloaderPlus
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void runButton_Click(object sender, EventArgs e)
        {
            if(sheetsListBox.CheckedItems.Count == 0)
            {
                MessageBox.Show("Please check at least one sheet to download images to.", "Please select sheet(s)");
                return;
            }

            if (MessageBox.Show("The downloaded images will be saved to this workbook. Do you want to continue?", "Download Images to Workbook?", MessageBoxButtons.YesNo) == DialogResult.No)
                return;

            Cursor.Current = Cursors.WaitCursor;

            statusLabel.Text = "Reading sheet cells...";
            ExcelHelper.DownloadImages(
                filenameTextBox.Text, sheetsListBox.CheckedItems.Cast<string>().ToArray(),
                downloadWriteProgressBar, statusLabel
            );

            if(MessageBox.Show("Your excel workbook now contains your images. Would you like to open it now?", "Done", MessageBoxButtons.YesNo) == DialogResult.Yes)
                Process.Start(filenameTextBox.Text);

            downloadWriteProgressBar.Value = 0;
            sheetsListBox.Items.Clear();
            statusLabel.Text = "";
            filenameTextBox.Text = "";
        }

        private void selectFileButton_Click(object sender, EventArgs e)
        {
            if (fileDialog.ShowDialog() != DialogResult.OK)
                return;

            filenameTextBox.Text = "";
            Cursor.Current = Cursors.WaitCursor;
            statusLabel.Text = "Opening workbook and reading sheets...";
            string[] sheetNames = ExcelHelper.ReadSheetNames(fileDialog.FileName);

            if(sheetNames == null)
            {
                Cursor.Current = Cursors.Default;
                statusLabel.Text = "";
                MessageBox.Show("Your excel workbook could not be opened. Do you already have it open in Excel? If so, you need to close it first.", "Cannot open");
                return;
            }

            filenameTextBox.Text = fileDialog.FileName;
            statusLabel.Text = "";
            sheetsListBox.Items.Clear();
            foreach(string sheetName in sheetNames)
            {
                sheetsListBox.Items.Add(sheetName, sheetNames.Length == 1 ? true : false);
            }
        }
    }
}
