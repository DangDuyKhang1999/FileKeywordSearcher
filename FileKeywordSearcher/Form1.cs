﻿using System.Diagnostics;
using System.IO;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Text;
using System.Windows.Forms;


namespace FileKeywordSearcher
{
    public partial class Form1 : Form
    {
        public enum FileExtension
        {
            Normal,
            CSV,
            Excel,
        }

        FileKeywordSearcher fileKeywordSearcher = null!;
        public Form1()
        {
            InitializeComponent();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            StartPosition = FormStartPosition.CenterScreen;
            Resize += Form1_SizeChanged;
        }

        private void btnBrowser_Click(object sender, EventArgs e)
        {
            txtBrowser.Text = String.Empty;
            txtBrowser.ForeColor = Color.Black;
            FolderBrowserDialog folderBrowserDialog = new();

            DialogResult result = folderBrowserDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                string selectedFolderPath = folderBrowserDialog.SelectedPath;

                txtBrowser.Text = selectedFolderPath;
            }
        }

        private void btnStartSearch_Click_1(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(txtBrowser.Text))
            {
                MessageBox.Show("Please select the directory for inspection!!!", "Error!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                btnBrowser.Focus();
                return;
            }

            if (!Directory.Exists(txtBrowser.Text))
            {
                MessageBox.Show("The directory is not valid!!!", "Error!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                btnBrowser.Focus();
                return;
            }

            fileKeywordSearcher = new FileKeywordSearcher(txtBrowser.Text, txtKeyWord.Text);

            if (!InitializeTableLayoutResult())
            {
                //MessageBox.Show("Not find {txtBrowser.text} int", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private bool InitializeTableLayoutResult()
        {
            if (fileKeywordSearcher == null)
            {
                return false;
            }

            bool bIsResult = false;
            List<FileItem> fileItems = fileKeywordSearcher.GetFileItems();
            int i = 0;
            tableLayoutPanel.RowStyles.Clear();
            tableLayoutPanel.Controls.Clear();

            if (fileItems.Count != 0)
            {
                int rtItemWidth = 0;
                if (fileItems.Count < 6)
                {
                    rtItemWidth = tableLayoutPanel.ClientSize.Width - 88;
                }
                else
                {
                    rtItemWidth = tableLayoutPanel.ClientSize.Width - 106;
                }

                bIsResult = true;

                foreach (FileItem fileItem in fileItems)
                {
                    TableLayoutPanel itemPanel = new()
                    {
                        Size = new Size(tableLayoutPanel.ClientSize.Width, 60),
                        ColumnCount = 2,
                        BackColor = Color.FromArgb(140, 194, 183)
                    };

                    //RichTextBox
                    string linecode = "";
                    if (fileItem.m_fileExtension == FileExtension.Normal)
                    {
                        linecode = $"   Line: {fileItem.m_strLineMapping}";
                    }
                    else if (fileItem.m_fileExtension == FileExtension.CSV)
                    {
                        linecode = $"   Cell: {fileItem.m_strLineMapping}";
                    }
                    else if (fileItem.m_fileExtension == FileExtension.Excel)
                    {
                        linecode = $"   Cell: {fileItem.m_strLineMapping}";
                    }

                    RichTextBox rtItem = new()
                    {
                        Size = new Size(rtItemWidth, 54),
                        Location = new Point(0, 0),
                        BorderStyle = BorderStyle.None
                    };

                    Font fontPath = new(rtItem.Font, FontStyle.Bold);
                    Font fontLine = new(rtItem.Font, FontStyle.Italic);

                    // fontLine for path
                    rtItem.SelectionStart = rtItem.TextLength;
                    rtItem.SelectionLength = 0;
                    rtItem.SelectionFont = fontPath;
                    rtItem.SelectionColor = Color.FromArgb(162, 87, 114);
                    rtItem.SelectedText = fileItem.m_strFileName + Environment.NewLine;

                    // fontLine for line
                    rtItem.SelectionStart = rtItem.TextLength;
                    rtItem.SelectionLength = 0;
                    rtItem.SelectionFont = fontLine;
                    rtItem.SelectionColor = Color.Black;
                    rtItem.SelectedText = linecode;

                    //Button
                    Button button = new()
                    {
                        Text = "Open" + Environment.NewLine + "Folder",
                        Size = new Size(70, 54),
                        Location = new Point(599, 0),
                        TextAlign = ContentAlignment.MiddleCenter,
                        ForeColor = Color.Black,
                        BackColor = Color.White
                    };
                    button.FlatAppearance.BorderColor = Color.FromArgb(137, 190, 179);
                    button.FlatStyle = FlatStyle.Flat;

                    button.Click += (sender, e) =>
                    {
                        if (sender is not null)
                        {
                            ButtonOpen_Click(sender, e, fileItem.m_strFileName);
                        }
                    };

                    itemPanel.Controls.Add(rtItem, 0, 0);
                    itemPanel.Controls.Add(button, 1, 0);

                    tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent));

                    tableLayoutPanel.Controls.Add(itemPanel, 0, i);
                    i++;
                    fontPath.Dispose();
                    fontLine.Dispose();
                }
            }
            return bIsResult;
        }

        private static void ButtonOpen_Click(object sender, EventArgs e, string filePath)
        {
            if (sender is null)
            {
                throw new ArgumentNullException(nameof(sender));
            }

            if (e is null)
            {
                throw new ArgumentNullException(nameof(e));
            }

            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentException($"'{nameof(filePath)}' cannot be null or empty.", nameof(filePath));
            }

            string directoryPath = Path.GetDirectoryName(filePath) ?? string.Empty;

            if (!string.IsNullOrEmpty(directoryPath) && Directory.Exists(directoryPath))
            {
                // Get file name from path
                string fileName = Path.GetFileName(filePath);

                // Open the folder and highlight the file
                _ = Process.Start("explorer.exe", "/select, " + Path.Combine(directoryPath, fileName));
            }
            else
            {
                MessageBox.Show("The folder does not exist!");
            }

        }

        private void UpdateControlSizesAndLocations()
        {
            txtBrowser.Width = ClientRectangle.Width - (ClientRectangle.Width - btnBrowser.Location.X + 5);
            txtBrowser.Height = btnBrowser.Height;

            Point newLocation = txtBrowser.Location;
            newLocation.Y = btnBrowser.Location.Y;
            txtBrowser.Location = newLocation;

            tableLayoutPanel.Width = ClientRectangle.Width;
            tableLayoutPanel.Height = ClientRectangle.Height - (ClientRectangle.Height - btnStartSearch.Location.Y - btnStartSearch.Height - 6);

            foreach (Control control in tableLayoutPanel.Controls)
            {
                if (control is Panel panel)
                {
                    panel.Width = tableLayoutPanel.ClientSize.Width;

                    RichTextBox rtb = panel.Controls.OfType<RichTextBox>().FirstOrDefault() ?? new RichTextBox();
                    Button button = panel.Controls.OfType<Button>().FirstOrDefault() ?? new Button();

                    if (rtb != null && button != null)
                    {
                        int newRtbWidth = panel.Width - button.Width - 12;

                        if (newRtbWidth > 0)
                        {
                            rtb.Width = newRtbWidth;
                        }
                    }
                }
            }
        }

        private void txtBrowser_Leave(object sender, EventArgs e)
        {
            if (txtBrowser.Text == String.Empty)
            {
                txtBrowser.Text = "Please select the directory for inspection!!!";
                txtBrowser.ForeColor = Color.Red;
            }
        }

        private void txtBrowser_Enter(object sender, EventArgs e)
        {
            if (txtBrowser.Text == "Please select the directory for inspection!!!")
            {
                txtBrowser.Text = String.Empty;
                txtBrowser.ForeColor = Color.Black;
            }
        }

        private void txtKeyWord_Enter(object sender, EventArgs e)
        {
            if (txtKeyWord.Text == "Enter the search keyword!!!")
            {
                txtKeyWord.Text = String.Empty;
                txtKeyWord.ForeColor = Color.Black;
            }
        }

        private void txtKeyWord_Leave(object sender, EventArgs e)
        {
            if (txtKeyWord.Text == String.Empty)
            {
                txtKeyWord.Text = "Enter the search keyword!!!";
                txtKeyWord.ForeColor = Color.Red;
            }
        }

        private void Form1_SizeChanged(object? sender, EventArgs e)
        {
            UpdateControlSizesAndLocations();
        }


    }
}