﻿using System.Diagnostics;
using System.IO;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Threading.Tasks;


namespace FileKeywordSearcher
{
    public partial class Form1 : Form
    {
        public enum FileExtension
        {
            Normal,
            CSV,
            Excel,
            Excel_Old,
            Word,
            Word_Old,
            Word_RTF,
            PowerPoint,
            PowerPoint_old,
            PDF,
            IgnoredExtension,
        }

        private FileKeywordSearcher fileKeywordSearcher = null!;
        private ProgressBar? progressBar1 = null!;
        private Label? txtProgressPercent = null!;
        private Label? txtProgressDetail = null!;
        private Label? txtProgressFileHasKeyWord = null!;
        public Form1()
        {
            InitializeComponent();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            StartPosition = FormStartPosition.CenterScreen;
            Resize += Form1_SizeChanged;
            SizeChanged += (sender, e) => { UpdateProgressBarWidth(); UpdateProgressBarPosition(); UpdateProgressBarFont(); };
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

        private async void btnStartSearch_Click_1(object sender, EventArgs e)
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
            if (!fileKeywordSearcher.getTotalFiles())
            {
                return;
            }
            ControlsStatus(false);
            if (progressBar1 == null)
            {
                tableLayoutPanel.Controls.Clear();
                InitializeProgressBarAndFileProcess();
            }
            // Show and start ProgressBar
            if (progressBar1 != null)
            {
                progressBar1.Visible = true;
                progressBar1.Value = 0;
            }
            UpdateControlSizesAndLocations();
            UpdateProgressBarWidth();
            UpdateProgressBarPosition();
            UpdateProgressBarFont();
            // Asynchronously call ProcessFiles method
            await Task.Run(() => fileKeywordSearcher.HasKeyWord(txtBrowser.Text));
        }

        private bool InitializeTableLayoutResult()
        {
            if (fileKeywordSearcher == null)
            {
                return false;
            }
            bool bIsResult = false;
            List<FileItem> fileItems = fileKeywordSearcher.GetFileItems();
            if (fileItems.Count == 0)
            {
                // Clear existing controls in the TableLayoutPanel
                tableLayoutPanel.Controls.Clear();

                // Add a Label with the message
                Label labelNoResult = new Label();
                labelNoResult.Text = "Keyword not found in directory!!!";
                labelNoResult.AutoSize = true;
                labelNoResult.Dock = DockStyle.Fill;
                labelNoResult.TextAlign = ContentAlignment.MiddleCenter;

                // Set the text color to red and make it bold
                labelNoResult.ForeColor = Color.Red;
                labelNoResult.Font = new Font(labelNoResult.Font, FontStyle.Bold);

                tableLayoutPanel.Controls.Add(labelNoResult, 0, 0);
                return false;
            }
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
                    switch (fileItem.m_fileExtension)
                    {
                        case FileExtension.Normal:
                            linecode = fileItem.m_bHasMultiKeyWord ? $"   Lines: {fileItem.m_strLineMapping}" : $"   Line: {fileItem.m_strLineMapping}";
                            break;
                        case FileExtension.CSV:
                            linecode = fileItem.m_bHasMultiKeyWord ? $"   Cells: {fileItem.m_strLineMapping}" : $"   Cell: {fileItem.m_strLineMapping}";
                            break;
                        case FileExtension.Excel:
                        case FileExtension.Excel_Old:
                            linecode = $"   {fileItem.m_strLineMapping}";
                            break;
                        case FileExtension.PDF:
                            linecode = fileItem.m_bHasMultiKeyWord ? $"   Pages: {fileItem.m_strLineMapping}" : $"   Page: {fileItem.m_strLineMapping}";
                            break;
                        case FileExtension.Word:
                        case FileExtension.Word_RTF:
                        case FileExtension.Word_Old:
                        case FileExtension.PowerPoint:
                            linecode = $"   Keyword detected in the file";
                            break;
                        default:
                            linecode = $"   {fileItem.m_strLineMapping}";
                            break;
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
            UpdateProgressBarWidth();
            UpdateProgressBarPosition();
            UpdateProgressBarFont();
        }

        // ProcessBar
        private void InitializeProgressBarAndFileProcess()
        {
            // Initialize ProgressBar
            progressBar1 = new ProgressBar
            {
                Minimum = 0,
                Maximum = 100,
                Step = 1,
                Visible = false,
                Height = ClientRectangle.Height / 15,
            };

            // Initialize TextBox Progress Precent
            txtProgressPercent = new Label
            {
                TextAlign = ContentAlignment.MiddleCenter,
                BorderStyle = BorderStyle.None,
                Height = progressBar1.Height,
                Width = progressBar1.Width,
                BackColor = Color.FromArgb(190, 217, 217),
            };
            // Initialize TextBox Progress Detail
            txtProgressDetail = new Label
            {
                TextAlign = ContentAlignment.TopLeft,
                BorderStyle = BorderStyle.None,
                Height = progressBar1.Height,
                Width = progressBar1.Width,
                BackColor = Color.FromArgb(190, 217, 217),
            };

            // Initialize TextBox File Path
            txtProgressFileHasKeyWord = new Label
            {
                TextAlign = ContentAlignment.TopRight,
                BorderStyle = BorderStyle.None,
                Height = progressBar1.Height,
                Width = progressBar1.Width,
                BackColor = Color.FromArgb(190, 217, 217),
            };

            //Position
            UpdateProgressBarWidth();
            UpdateProgressBarPosition();
            UpdateProgressBarFont();

            // Add controls to Form
            this.Controls.Add(progressBar1);
            this.Controls.Add(txtProgressPercent);
            this.Controls.Add(txtProgressDetail);
            this.Controls.Add(txtProgressFileHasKeyWord);

            // Bring ProgressBar to front
            progressBar1.BringToFront();
            txtProgressPercent.BringToFront();
            txtProgressDetail.BringToFront();
            txtProgressFileHasKeyWord.BringToFront();

            // Initialize FileProcess instance and subscribe to ProgressChanged event
            fileKeywordSearcher.ProgressChanged += FileProcessor_ProgressChanged;
        }

        private void UpdateProgressBarWidth()
        {
            if (progressBar1 != null && txtProgressPercent != null && txtProgressDetail != null && txtProgressFileHasKeyWord != null)
            {
                progressBar1.Width = ClientRectangle.Width - 50;
                progressBar1.Height = ClientRectangle.Height / 15;

                txtProgressPercent.Width = progressBar1.Width;
                txtProgressPercent.Height = progressBar1.Height;

                txtProgressDetail.Width = progressBar1.Width / 2;
                txtProgressDetail.Height = progressBar1.Height + 20;

                txtProgressFileHasKeyWord.Width = progressBar1.Width / 2;
                txtProgressFileHasKeyWord.Height = progressBar1.Height + 20;
            }
        }
        private void UpdateProgressBarPosition()
        {
            if (progressBar1 != null && txtProgressPercent != null && txtProgressDetail != null && txtProgressFileHasKeyWord != null)
            {
                int progressBarHeight = progressBar1.Height;
                int progressBarX = (ClientRectangle.Width - progressBar1.Width) / 2;
                int progressBarY = (ClientRectangle.Height - progressBarHeight) / 2;

                progressBar1.Location = new Point(progressBarX, progressBarY);
                txtProgressPercent.Location = new Point(progressBarX, progressBarY - txtProgressPercent.Height - 10);
                txtProgressDetail.Location = new Point(progressBarX, progressBarY + progressBar1.Height);
                txtProgressFileHasKeyWord.Location = new Point(progressBarX + progressBar1.Width / 2, progressBarY + progressBar1.Height);
            }
        }

        private void UpdateProgressBarFont()
        {
            if (progressBar1 != null && txtProgressPercent != null && txtProgressDetail != null && txtProgressFileHasKeyWord != null)
            {
                Font font = new Font("Segoe UI", progressBar1.Height / 2, FontStyle.Bold, GraphicsUnit.Point);
                txtProgressDetail.Font = font;
                txtProgressPercent.Font = font;
                txtProgressFileHasKeyWord.Font = font;
            }
        }


        private void FileProcessor_ProgressChanged(object sender, (int percent, int iFileCount, int iTotalFileCount, int iFileHasKeyWord) e)
        {
            // Handle ProgressChanged event from fileProcessor
            _ = this.Invoke((MethodInvoker)delegate ()
            {
                progressBar1.Value = e.percent;
                progressBar1.Refresh(); // Ensure ProgressBar updates visually
                txtProgressPercent.Text = e.percent.ToString() + "%";
                txtProgressDetail.Text = e.iFileCount.ToString() + "/" + e.iTotalFileCount.ToString();
                txtProgressFileHasKeyWord.Text = "Files containing keyword: " + e.iFileHasKeyWord.ToString();
                // Check if progress is complete (100%)
                if (e.percent >= 100)
                {
                    // Remove ProgressBar from Form
                    System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
                    timer.Interval = 500;
                    timer.Tick += (s, e) =>
                    {
                        timer.Stop();

                        progressBar1.Visible = false;
                        progressBar1 = null;

                        txtProgressPercent.Visible = false;
                        txtProgressPercent = null;

                        txtProgressDetail.Visible = false;
                        txtProgressDetail = null;

                        txtProgressFileHasKeyWord.Visible = false;
                        txtProgressFileHasKeyWord = null;

                        // Remove ProgressBar from Form
                        this.Controls.Remove(progressBar1);
                        this.Controls.Remove(txtProgressPercent);
                        this.Controls.Remove(txtProgressDetail);
                        this.Controls.Remove(txtProgressFileHasKeyWord);
                        InitializeTableLayoutResult();
                        ControlsStatus(true);

                        // Optionally unsubscribe from ProgressChanged event to prevent further updates
                        fileKeywordSearcher.ProgressChanged -= FileProcessor_ProgressChanged;
                    };
                    timer.Start();
                }
            });
        }

        private void ControlsStatus(bool isEnable)
        {
            if (isEnable)
            {
                txtKeyWord.Enabled = true;
                txtBrowser.Enabled = true;
                btnBrowser.Enabled = true;
                btnStartSearch.Enabled = true;
                txtKeyWord.BackColor = Color.FromArgb(137, 190, 179);
                txtBrowser.BackColor = Color.FromArgb(137, 190, 179);
                btnBrowser.BackColor = Color.FromArgb(137, 190, 179);
                btnStartSearch.BackColor = Color.FromArgb(137, 190, 179);
            }
            else
            {
                txtKeyWord.Enabled = false;
                txtBrowser.Enabled = false;
                btnBrowser.Enabled = false;
                btnStartSearch.Enabled = false;
                txtKeyWord.BackColor = Color.LightGray;
                txtBrowser.BackColor = Color.LightGray;
                btnBrowser.BackColor = Color.LightGray;
                btnStartSearch.BackColor = Color.LightGray;
            }
        }

    }
}