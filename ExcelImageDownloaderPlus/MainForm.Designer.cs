
namespace ExcelImageDownloaderPlus
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);

            ExcelHelper.CleanupOnExit();
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.runButton = new System.Windows.Forms.Button();
            this.fileDialog = new System.Windows.Forms.OpenFileDialog();
            this.filenameTextBox = new System.Windows.Forms.TextBox();
            this.selectFileButton = new System.Windows.Forms.Button();
            this.sheetsListBox = new System.Windows.Forms.CheckedListBox();
            this.sheetsListLabel = new System.Windows.Forms.Label();
            this.downloadWriteProgressBar = new System.Windows.Forms.ProgressBar();
            this.statusLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // runButton
            // 
            this.runButton.Location = new System.Drawing.Point(551, 404);
            this.runButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.runButton.Name = "runButton";
            this.runButton.Size = new System.Drawing.Size(235, 42);
            this.runButton.TabIndex = 0;
            this.runButton.Text = "Run!";
            this.runButton.UseVisualStyleBackColor = true;
            this.runButton.Click += new System.EventHandler(this.runButton_Click);
            // 
            // fileDialog
            // 
            this.fileDialog.Filter = "Excel xlsx files|*.xlsx";
            // 
            // filenameTextBox
            // 
            this.filenameTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.filenameTextBox.Location = new System.Drawing.Point(191, 26);
            this.filenameTextBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.filenameTextBox.Name = "filenameTextBox";
            this.filenameTextBox.ReadOnly = true;
            this.filenameTextBox.Size = new System.Drawing.Size(595, 26);
            this.filenameTextBox.TabIndex = 1;
            // 
            // selectFileButton
            // 
            this.selectFileButton.Location = new System.Drawing.Point(12, 26);
            this.selectFileButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.selectFileButton.Name = "selectFileButton";
            this.selectFileButton.Size = new System.Drawing.Size(172, 32);
            this.selectFileButton.TabIndex = 2;
            this.selectFileButton.Text = "Select XLSX...";
            this.selectFileButton.UseVisualStyleBackColor = true;
            this.selectFileButton.Click += new System.EventHandler(this.selectFileButton_Click);
            // 
            // sheetsListBox
            // 
            this.sheetsListBox.FormattingEnabled = true;
            this.sheetsListBox.Location = new System.Drawing.Point(14, 96);
            this.sheetsListBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.sheetsListBox.Name = "sheetsListBox";
            this.sheetsListBox.Size = new System.Drawing.Size(772, 193);
            this.sheetsListBox.TabIndex = 3;
            // 
            // sheetsListLabel
            // 
            this.sheetsListLabel.AutoSize = true;
            this.sheetsListLabel.Location = new System.Drawing.Point(9, 72);
            this.sheetsListLabel.Name = "sheetsListLabel";
            this.sheetsListLabel.Size = new System.Drawing.Size(238, 20);
            this.sheetsListLabel.TabIndex = 4;
            this.sheetsListLabel.Text = "Sheets to download images to:";
            // 
            // downloadWriteProgressBar
            // 
            this.downloadWriteProgressBar.Location = new System.Drawing.Point(12, 340);
            this.downloadWriteProgressBar.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.downloadWriteProgressBar.Name = "downloadWriteProgressBar";
            this.downloadWriteProgressBar.Size = new System.Drawing.Size(774, 58);
            this.downloadWriteProgressBar.TabIndex = 5;
            // 
            // statusLabel
            // 
            this.statusLabel.AutoSize = true;
            this.statusLabel.Location = new System.Drawing.Point(9, 315);
            this.statusLabel.MaximumSize = new System.Drawing.Size(774, 20);
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(0, 20);
            this.statusLabel.TabIndex = 6;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 460);
            this.Controls.Add(this.statusLabel);
            this.Controls.Add(this.downloadWriteProgressBar);
            this.Controls.Add(this.sheetsListLabel);
            this.Controls.Add(this.sheetsListBox);
            this.Controls.Add(this.selectFileButton);
            this.Controls.Add(this.filenameTextBox);
            this.Controls.Add(this.runButton);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel Image Downloader Plus";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button runButton;
        private System.Windows.Forms.OpenFileDialog fileDialog;
        private System.Windows.Forms.TextBox filenameTextBox;
        private System.Windows.Forms.Button selectFileButton;
        private System.Windows.Forms.CheckedListBox sheetsListBox;
        private System.Windows.Forms.Label sheetsListLabel;
        private System.Windows.Forms.ProgressBar downloadWriteProgressBar;
        private System.Windows.Forms.Label statusLabel;
    }
}

