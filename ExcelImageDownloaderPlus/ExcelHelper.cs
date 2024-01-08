using System;
using System.Collections.Generic;

using System.Net;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;

using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace ExcelImageDownloaderPlus {
    class ExcelHelper {
        private static Microsoft.Office.Interop.Excel.Application excelInstance = new Microsoft.Office.Interop.Excel.Application();
        private static Workbook book = null;

        public static string[] ReadSheetNames(string filename) {
            // Open the workbook
            if (book != null)
                book.Close(false);

            book = excelInstance.Workbooks.Open(filename, false, false, IgnoreReadOnlyRecommended: true);
            if (book == null)
            {
                return null;
            } else if (book.ReadOnly)
            {
                book.Close(false);
                book = null;
                return null;
            }

            List<string> sheetNames = new List<string>();
            foreach(Worksheet sheet in book.Sheets)
            {
                if (sheet.Visible == XlSheetVisibility.xlSheetVisible)
                {
                    sheetNames.Add(sheet.Name);
                }
            }

            return sheetNames.ToArray();
        }

        private static bool IsCellValueImageString(string value)
        {
            if (
                value.Contains("https://s3.amazonaws.com/com.modaoperandi.cdn.assets.v2/images")
                || value.Contains("https://cdn.modaoperandi.com/img/images")
                || value.Contains("https://cdn.modaoperandi.com/assets/images")
                || value.Contains("https://cdn.modaoperandi.com/img/uploads")
                || value.Contains("https://cdn.modaoperandi.com/assets/uploads")
                || value.Contains("https://www.modaoperandi.com/assets/images")
            )
                return true;

            return false;
        }

        public static void DownloadAllImages(List<Range> cellList, ProgressBar progressBar)
        {
            using (var webClient = new WebClient())
            {
                foreach (Range cell in cellList)
                {
                    string filename = System.IO.Directory.GetCurrentDirectory() + "\\" + cell.Column.ToString() + "-" + cell.Row.ToString() + "-temp-jpg.jpg";
                
                    try
                    {
                        // Set these on each download request. The CDN may refuse the request otherwise.
                        webClient.Headers.Add("Content-type", "image/jpeg");
                        webClient.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:59.0) Gecko/20100101 Firefox/59.0");

                        byte[] imageBytes = webClient.DownloadData((string)cell.Value2);

                        using (MemoryStream ms = new MemoryStream(imageBytes))
                        {
                            Image jpg = Image.FromStream(ms);
                            jpg.Save(filename, ImageFormat.Jpeg);
                        }
                    } catch(WebException)
                    {
                        // Keep going, this is a missing image.
                    }

                    progressBar.Value++;
                }
            }
        }

        public static void InsertImagesToCells(List<Range> cellList, ProgressBar progressBar)
        {
            foreach (Range cell in cellList)
            {
                string filename = System.IO.Directory.GetCurrentDirectory() + "\\" + cell.Column.ToString() + "-" + cell.Row.ToString() + "-temp-jpg.jpg";
                if(!File.Exists(filename))
                {
                    cell.Value2 = "Image unavailable";
                    continue;
                }

                // Insert into sheet
                try
                {
                    Microsoft.Office.Interop.Excel.Shape picture = cell.Worksheet.Shapes.AddPicture(
                        filename, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0, -1, -1
                    );

                    picture.Height = 100;
                    cell.RowHeight = picture.Height + 20;
                    cell.ColumnWidth = picture.Width * (54.29 / 288);
                    picture.Left = (float)cell.Left;
                    picture.Top = (float)cell.Top;
                    picture.Placement = XlPlacement.xlMoveAndSize;

                    cell.Value = "";
                } catch(Exception)
                {
                    cell.Value = "Image unavailable";
                }

                progressBar.Value++;

                // Delete temp jpg
                File.Delete(filename);
            }
        }

        public static void DownloadImages(string filename, string[] sheetNames, ProgressBar progressBar, System.Windows.Forms.Label statusLabel)
        {
            foreach (string sheetName in sheetNames)
            {
                // Get the sheet
                Worksheet sheet = (Worksheet)book.Sheets[sheetName];
                List<Range> cellList = new List<Range>();

                // Go through each cell in the full sheet range
                Range fullRange = sheet.UsedRange;
                statusLabel.Text = "Reading " + sheetName + " cells...";

                foreach (Range col in fullRange.Columns)
                {
                    // Get all cells that need a downloaded image
                    foreach(Range cell in col.Rows)
                    {
                        statusLabel.Text = "Reading cell " + cell.Column.ToString() + ":" + cell.Row.ToString() + "...";
                        dynamic cellValue = cell.Value2;
                        // Check if the cell is a string and image URL
                        if (cellValue == null)
                            continue;

                        //string cellType = cellValue.GetType().ToString();
                        try
                        {
                            string stringValue = (string)cellValue;
                            if (IsCellValueImageString(stringValue))
                                cellList.Add(cell);
                        }
                        catch(InvalidCastException)
                        {
                            // Not casted to a string, continue
                        }
                        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                        {
                            // Not casted to a string, continue
                        }
                    }
                }

                progressBar.Value = 0;
                progressBar.Maximum = cellList.Count * 2;
                statusLabel.Text = "Downloading " + cellList.Count + " images...";

                // Download every image.
                DownloadAllImages(cellList, progressBar);

                statusLabel.Text = "Inserting " + cellList.Count + " images into sheet...";
                
                // Insert all images and delete temp files.
                InsertImagesToCells(cellList, progressBar);
            }

            statusLabel.Text = "Saving workbook and closing...";

            book.Save();
            book.Close(false);
            book = null;
            statusLabel.Text = "Complete.";
        }

        public static void CleanupOnExit()
        {
            if (book != null)
                book.Close(false);
            excelInstance.Quit();
        }
    }
}
