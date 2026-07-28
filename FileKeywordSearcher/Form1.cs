using System.Diagnostics;
using System.IO;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Threading;
using System.Runtime.InteropServices;
using Org.BouncyCastle.Crypto;
using Microsoft.WindowsAPICodePack.Taskbar;


namespace FileKeywordSearcher
{
    public partial class Form1 : Form
    {
        private CancellationTokenSource cancellationTokenSource;
        private FileKeywordSearcher fileKeywordSearcher = null!;
        private ModernProgressBar? progressBar1 = null!;
        private Label? txtProgressPercent = null!;
        private Label? txtProgressDetail = null!;
        private Label? txtProgressFileHasKeyWord = null!;
        private Label? txtProgressCurrentFile = null!;
        private const int InitialVisibleResults = 5;
        private int? _resultRowHeight;
        private int _resultPage;
        private bool _isSearchRunning;
        private bool _stopRequested;
        private const int DwmwaBorderColor = 34;
        private const int DwmwaCaptionColor = 35;
        private const int DwmwaTextColor = 36;

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int valueSize);

        public Form1()
        {
            InitializeComponent();
            HandleCreated += (_, _) => ApplyPastelTitleBar(this);
            cancellationTokenSource = new CancellationTokenSource();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            StartPosition = FormStartPosition.CenterScreen;
            Resize += Form1_SizeChanged;
            ResizeEnd += (_, _) =>
            {
                if (!_isSearchRunning && tableLayoutPanel.Visible && fileKeywordSearcher != null)
                    InitializeTableLayoutResult();
            };
            SizeChanged += (sender, e) => { UpdateProgressBarWidth(); UpdateProgressBarPosition(); UpdateProgressBarFont(); };
        }

        private void btnBrowser_Click(object sender, EventArgs e)
        {
            txtBrowser.Text = String.Empty;
            txtBrowser.ForeColor = Color.FromArgb(43, 74, 55);
            FolderBrowserDialog folderBrowserDialog = new();

            DialogResult result = folderBrowserDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                string selectedFolderPath = folderBrowserDialog.SelectedPath;

                txtBrowser.Text = selectedFolderPath;
            }
        }
        private void FileProcessor_ProgressChanged(object? sender, (int percent, int iFileCount, int iTotalFileCount, int iFileHasKeyWord, string strCurrentFile) e)
        {
            // Ensure UI updates are invoked on the UI thread
            _ = this.Invoke((MethodInvoker)delegate ()
            {
                // Update ProgressBar
                if (progressBar1 != null)
                {
                    progressBar1.IsIndeterminate = false;
                    progressBar1.Value = e.percent;
                    progressBar1.Refresh(); // Ensure ProgressBar updates visually
                    TaskbarManager.Instance.SetProgressValue(e.percent, 100);
                    TaskbarManager.Instance.SetProgressState(TaskbarProgressBarState.Normal);
                }

                // Update progress text details
                if (txtProgressPercent != null)
                {
                    txtProgressPercent.Text = e.percent + "%";
                }

                if (txtProgressDetail != null)
                {
                    txtProgressDetail.Text = $"{e.iFileCount:N0}/{e.iTotalFileCount:N0}";
                }

                if (txtProgressFileHasKeyWord != null)
                {
                    txtProgressFileHasKeyWord.Text = $"Files containing keyword: {e.iFileHasKeyWord}";
                }

                if (txtProgressCurrentFile != null)
                {
                    txtProgressCurrentFile.Text = e.strCurrentFile;
                    txtProgressCurrentFile.Height = txtProgressCurrentFile.GetPreferredSize(new Size(txtProgressCurrentFile.Width, int.MaxValue)).Height;
                }

                if (e.percent >= 100)
                {
                    TaskbarManager.Instance.SetProgressState(TaskbarProgressBarState.NoProgress);
                }
            });
        }

        private async void btnStartSearch_Click_1(object sender, EventArgs e)
        {
            if (_isSearchRunning)
            {
                if (_stopRequested) return;
                _stopRequested = true;
                btnStartSearch.Enabled = false;
                btnStartSearch.Text = "Stopping…";
                cancellationTokenSource.Cancel();
                TaskbarManager.Instance.SetProgressState(TaskbarProgressBarState.Paused);
                return;
            }

            if (txtBrowser.Text == "Please select the directory for searching!!!")
            {
                MessageBox.Show("Please select the directory for searching!!!", "Error!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                btnBrowser.Focus();
                return;
            }

            if (!Directory.Exists(txtBrowser.Text))
            {
                MessageBox.Show("The directory is not valid!!!", "Error!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                btnBrowser.Focus();
                return;
            }

            if (txtKeyWord.Text == "Enter the search keyword!!!")
            {
                MessageBox.Show("Please enter the keyword for the search!!!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtKeyWord.Focus();
                return;
            }

            _isSearchRunning = true;
            _stopRequested = false;
            cancellationTokenSource = new CancellationTokenSource();
            fileKeywordSearcher = new FileKeywordSearcher(txtBrowser.Text, txtKeyWord.Text, labelWithCheckBoxList.m_SelectedItems);
            _resultPage = 0;
            ControlsStatus(false);
            emptyStatePanel.Visible = false;
            tableLayoutPanel.Visible = false;
            tableLayoutPanel.AutoScroll = false;
            resultsPagerHost.Visible = false;

            try
            {
                if (progressBar1 == null)
                {
                    ClearResultControls();
                    InitializeProgressBarAndFileProcess();
                }
                if (progressBar1 != null)
                {
                    progressBar1.Visible = true;
                    progressBar1.IsIndeterminate = true;
                }
                if (txtProgressPercent != null) txtProgressPercent.Text = "Indexing files…";
                if (txtProgressDetail != null) txtProgressDetail.Text = "Preparing file list";
                if (txtProgressFileHasKeyWord != null) txtProgressFileHasKeyWord.Text = string.Empty;
                if (txtProgressCurrentFile != null) txtProgressCurrentFile.Text = txtBrowser.Text;
                UpdateProgressBarWidth();
                UpdateProgressBarPosition();
                UpdateProgressBarFont();

                bool hasFiles = await Task.Run(() => fileKeywordSearcher.getTotalFiles(cancellationTokenSource.Token));
                if (!hasFiles)
                {
                    ClearProgressBar();
                    MessageBox.Show("No supported files were found in this folder.", "Search complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (progressBar1 != null)
                {
                    progressBar1.IsIndeterminate = false;
                    progressBar1.Value = 0;
                }
                if (txtProgressPercent != null) txtProgressPercent.Text = "0%";
                UpdateControlSizesAndLocations();
                UpdateProgressBarWidth();
                UpdateProgressBarPosition();
                UpdateProgressBarFont();

                await fileKeywordSearcher.HasKeyWord(cancellationTokenSource.Token);
                if (!cancellationTokenSource.IsCancellationRequested)
                {
                    ClearProgressBar();
                    ShowSkippedFilesDialog(fileKeywordSearcher.GetSkippedFiles());
                }
            }
            catch (OperationCanceledException)
            {
                // Cancellation is an expected result of the Stop button.
            }
            finally
            {
                _isSearchRunning = false;
                _stopRequested = false;
                btnStartSearch.Enabled = true;
                TaskbarManager.Instance.SetProgressState(TaskbarProgressBarState.NoProgress);
                if (cancellationTokenSource.IsCancellationRequested && progressBar1 != null)
                    ClearProgressBar();
                ControlsStatus(true);
                btnStartSearch.Text = "Search";
            }
        }

        private bool InitializeTableLayoutResult()
        {
            tableLayoutPanel.AutoScroll = false;
            if (fileKeywordSearcher == null)
            {
                return false;
            }
            bool bIsResult = false;
            List<FileItem> fileItems = fileKeywordSearcher.GetFileItems();
            if (fileItems.Count == 0)
            {
                resultsPagerHost.Visible = false;
                resultsContentLayout.RowStyles[1].Height = 0F;
                // Clear existing controls in the TableLayoutPanel
                ClearResultControls();
                tableLayoutPanel.RowStyles.Clear();
                tableLayoutPanel.RowCount = 1;
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

                // Add a Label with the message
                Label labelNoResult = new Label();
                labelNoResult.Text = "No matching files found";
                labelNoResult.AutoSize = true;
                labelNoResult.Dock = DockStyle.Fill;
                labelNoResult.TextAlign = ContentAlignment.MiddleCenter;

                // Set the text color to red and make it bold
                labelNoResult.ForeColor = Color.FromArgb(83, 125, 96);
                labelNoResult.Font = new Font(labelNoResult.Font, FontStyle.Bold);

                tableLayoutPanel.Controls.Add(labelNoResult, 0, 0);
                emptyStatePanel.Visible = false;
                tableLayoutPanel.Visible = true;
                return false;
            }
            int i = 0;
            tableLayoutPanel.RowStyles.Clear();
            ClearResultControls();
            tableLayoutPanel.AutoScrollPosition = Point.Empty;

            if (fileItems.Count != 0)
            {
                bIsResult = true;
                int availableHeight = Math.Max(1, resultsContentLayout.ClientSize.Height - 58);
                _resultRowHeight ??= Math.Max(54, availableHeight / InitialVisibleResults);
                int resultRowHeight = _resultRowHeight.Value;
                int resultsPerPage = Math.Max(1, availableHeight / resultRowHeight);
                int pageCount = (int)Math.Ceiling(fileItems.Count / (double)resultsPerPage);
                _resultPage = Math.Clamp(_resultPage, 0, pageCount - 1);
                List<FileItem> visibleItems = fileItems
                    .Skip(_resultPage * resultsPerPage)
                    .Take(resultsPerPage)
                    .ToList();
                bool showPager = pageCount > 1;
                // One dedicated flexible row absorbs only the remainder after fitting
                // as many fixed-height records as possible. This prevents WinForms
                // from stretching the final record when the window grows.
                tableLayoutPanel.RowCount = visibleItems.Count + 1;

                foreach (FileItem fileItem in visibleItems)
                {
                    TableLayoutPanel itemPanel = new()
                    {
                        Dock = DockStyle.Fill,
                        ColumnCount = 2,
                        RowCount = 1,
                        BackColor = Color.FromArgb(235, 246, 238),
                        Margin = new Padding(2, 4, 8, 4),
                        Padding = new Padding(12, 8, 8, 8)
                    };
                    itemPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
                    itemPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 96F));
                    itemPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
                    itemPanel.Resize += (_, _) =>
                    {
                        if (itemPanel.Width <= 1 || itemPanel.Height <= 1) return;
                        using System.Drawing.Drawing2D.GraphicsPath rounded = RoundedPanel.RoundedPath(itemPanel.ClientRectangle, 12);
                        itemPanel.Region?.Dispose();
                        itemPanel.Region = new Region(rounded);
                    };

                    //RichTextBox
                    string linecode = "";
                    switch (fileItem.m_fileExtension)
                    {
                        case eFileExtension.Normal:
                            linecode = fileItem.m_bHasMultiKeyWord ? $"   Lines: {fileItem.m_strLineMapping}" : $"   Line: {fileItem.m_strLineMapping}";
                            break;
                        case eFileExtension.CSV:
                            linecode = fileItem.m_bHasMultiKeyWord ? $"   Cells: {fileItem.m_strLineMapping}" : $"   Cell: {fileItem.m_strLineMapping}";
                            break;
                        case eFileExtension.Excel:
                        case eFileExtension.Excel_Old:
                            linecode = $"   {fileItem.m_strLineMapping}";
                            break;
                        case eFileExtension.PDF:
                            linecode = fileItem.m_bHasMultiKeyWord ? $"   Pages: {fileItem.m_strLineMapping}" : $"   Page: {fileItem.m_strLineMapping}";
                            break;
                        case eFileExtension.Word:
                        case eFileExtension.Word_RTF:
                        case eFileExtension.Word_Old:
                        case eFileExtension.PowerPoint:
                        case eFileExtension.PowerPoint_old:
                            linecode = $"   Keyword detected in the file";
                            break;
                        default:
                            linecode = $"   {fileItem.m_strLineMapping}";
                            break;
                    }

                    TableLayoutPanel textPanel = new()
                    {
                        Dock = DockStyle.Fill,
                        BackColor = Color.FromArgb(235, 246, 238),
                        Margin = new Padding(0, 0, 10, 0),
                        ColumnCount = 1,
                        RowCount = 2
                    };
                    textPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
                    textPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 55F));
                    textPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 45F));

                    Label fileNameLabel = new()
                    {
                        AutoEllipsis = true,
                        Dock = DockStyle.Fill,
                        Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                        ForeColor = Color.FromArgb(55, 112, 73),
                        Text = fileItem.m_strFileName,
                        TextAlign = ContentAlignment.MiddleLeft,
                        UseMnemonic = false
                    };
                    Label locationLabel = new()
                    {
                        AutoEllipsis = true,
                        Dock = DockStyle.Fill,
                        Font = new Font("Segoe UI", 8.5F, FontStyle.Italic),
                        ForeColor = Color.FromArgb(100, 129, 109),
                        Text = linecode.TrimStart(),
                        TextAlign = ContentAlignment.MiddleLeft,
                        UseMnemonic = false
                    };
                    ToolTip fullPathTip = new() { InitialDelay = 350, AutoPopDelay = 12000 };
                    fullPathTip.SetToolTip(fileNameLabel, fileItem.m_strFileName);
                    textPanel.Tag = fullPathTip;
                    textPanel.Controls.Add(fileNameLabel, 0, 0);
                    textPanel.Controls.Add(locationLabel, 0, 1);

                    //Button
                    Button button = new()
                    {
                        Text = "Open",
                        Dock = DockStyle.Fill,
                        TextAlign = ContentAlignment.MiddleCenter,
                        ForeColor = Color.FromArgb(45, 91, 59),
                        BackColor = Color.FromArgb(211, 235, 218),
                        Cursor = Cursors.Hand,
                        Padding = Padding.Empty,
                        UseCompatibleTextRendering = false
                    };
                    button.FlatAppearance.BorderColor = Color.FromArgb(184, 214, 193);
                    button.FlatStyle = FlatStyle.Flat;

                    button.Click += (sender, e) =>
                    {
                        if (sender is not null)
                        {
                            ButtonOpen_Click(sender, e, fileItem.m_strFileName);
                        }
                    };

                    itemPanel.Controls.Add(textPanel, 0, 0);
                    itemPanel.Controls.Add(button, 1, 0);

                    tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, resultRowHeight));

                    tableLayoutPanel.Controls.Add(itemPanel, 0, i);
                    i++;
                }
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
                if (showPager)
                {
                    Panel pager = CreateResultsPager(pageCount, fileItems.Count, resultsPerPage);
                    resultsPagerHost.Controls.Clear();
                    resultsPagerHost.Controls.Add(pager);
                    resultsPagerHost.Visible = true;
                    resultsContentLayout.RowStyles[1].Height = 58F;
                }
                else
                {
                    resultsPagerHost.Visible = false;
                    resultsContentLayout.RowStyles[1].Height = 0F;
                }
            }
            emptyStatePanel.Visible = false;
            tableLayoutPanel.Visible = true;
            return bIsResult;
        }

        private Panel CreateResultsPager(int pageCount, int totalResults, int resultsPerPage)
        {
            TableLayoutPanel pager = new()
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(244, 250, 245),
                Padding = new Padding(6, 8, 6, 8),
                ColumnCount = 3,
                RowCount = 1,
                Margin = Padding.Empty
            };
            pager.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120F));
            pager.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            pager.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120F));
            pager.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            ModernButton previous = new()
            {
                Text = "‹  Previous",
                Enabled = _resultPage > 0,
                Dock = DockStyle.Fill,
                Margin = new Padding(4, 0, 8, 0),
                BackColor = Color.FromArgb(222, 240, 227),
                ForeColor = Color.FromArgb(55, 100, 69),
                BorderColor = Color.FromArgb(194, 220, 202),
                CornerRadius = 17
            };
            ModernButton next = new()
            {
                Text = "Next  ›",
                Enabled = _resultPage < pageCount - 1,
                Dock = DockStyle.Fill,
                Margin = new Padding(8, 0, 4, 0),
                BackColor = Color.FromArgb(222, 240, 227),
                ForeColor = Color.FromArgb(55, 100, 69),
                BorderColor = Color.FromArgb(194, 220, 202),
                CornerRadius = 17
            };
            Label pageInfo = new()
            {
                Dock = DockStyle.Fill,
                Margin = Padding.Empty,
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.FromArgb(83, 125, 96),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Text = $"Showing {_resultPage * resultsPerPage + 1:N0}–{Math.Min((_resultPage + 1) * resultsPerPage, totalResults):N0} of {totalResults:N0}   •   Page {_resultPage + 1}/{pageCount}"
            };
            previous.Click += (_, _) => { _resultPage--; tableLayoutPanel.AutoScrollPosition = Point.Empty; InitializeTableLayoutResult(); };
            next.Click += (_, _) => { _resultPage++; tableLayoutPanel.AutoScrollPosition = Point.Empty; InitializeTableLayoutResult(); };
            pager.Controls.Add(previous, 0, 0);
            pager.Controls.Add(pageInfo, 1, 0);
            pager.Controls.Add(next, 2, 0);
            return pager;
        }

        private void ClearResultControls()
        {
            resultsPagerHost.Visible = false;
            resultsContentLayout.RowStyles[1].Height = 0F;
            while (resultsPagerHost.Controls.Count > 0)
            {
                Control pagerControl = resultsPagerHost.Controls[0];
                resultsPagerHost.Controls.RemoveAt(0);
                pagerControl.Dispose();
            }
            while (tableLayoutPanel.Controls.Count > 0)
            {
                Control control = tableLayoutPanel.Controls[0];
                tableLayoutPanel.Controls.RemoveAt(0);
                control.Dispose();
            }
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
            BringToForntControl();
        }

        private void ShowSkippedFilesDialog(IReadOnlyList<(string FilePath, string Reason)> skippedFiles)
        {
            if (skippedFiles.Count == 0) return;

            using Form dialog = new()
            {
                Text = $"Skipped files ({skippedFiles.Count})",
                StartPosition = FormStartPosition.CenterParent,
                Size = new Size(760, 420),
                MinimumSize = new Size(560, 320),
                BackColor = Color.FromArgb(232, 243, 235),
                ForeColor = Color.FromArgb(43, 74, 55),
                Font = new Font("Segoe UI", 9F),
                ShowIcon = false
            };
            dialog.HandleCreated += (_, _) => ApplyPastelTitleBar(dialog);
            Label heading = new()
            {
                Dock = DockStyle.Top,
                Height = 58,
                Padding = new Padding(18, 12, 18, 4),
                Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                Text = "Some files could not be scanned and were skipped."
            };
            TextBox details = new()
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                BackColor = Color.FromArgb(244, 250, 245),
                ForeColor = Color.FromArgb(55, 100, 69),
                BorderStyle = BorderStyle.FixedSingle,
                Text = string.Join(Environment.NewLine + Environment.NewLine,
                    skippedFiles.Select((item, index) => $"{index + 1}. {item.FilePath}{Environment.NewLine}   Reason: {item.Reason}"))
            };
            ModernButton close = new()
            {
                Text = "Close",
                Dock = DockStyle.Right,
                Width = 110,
                BackColor = Color.FromArgb(137, 201, 158),
                ForeColor = Color.FromArgb(28, 73, 43),
                BorderColor = Color.FromArgb(116, 185, 139),
                CornerRadius = 10,
                DialogResult = DialogResult.OK,
                Margin = new Padding(0, 8, 18, 8)
            };
            Panel footer = new()
            {
                Dock = DockStyle.Bottom,
                Height = 58,
                Padding = new Padding(0, 10, 18, 10),
                BackColor = Color.FromArgb(232, 243, 235)
            };
            footer.Controls.Add(close);
            dialog.Controls.Add(details);
            dialog.Controls.Add(heading);
            dialog.Controls.Add(footer);
            dialog.AcceptButton = close;
            dialog.CancelButton = close;
            dialog.ShowDialog(this);
        }

        private static void ApplyPastelTitleBar(Form form)
        {
            if (!OperatingSystem.IsWindows()) return;
            try
            {
                int captionColor = ToColorRef(Color.FromArgb(224, 240, 228));
                int borderColor = ToColorRef(Color.FromArgb(190, 216, 198));
                int textColor = ToColorRef(Color.FromArgb(43, 74, 55));
                DwmSetWindowAttribute(form.Handle, DwmwaCaptionColor, ref captionColor, sizeof(int));
                DwmSetWindowAttribute(form.Handle, DwmwaBorderColor, ref borderColor, sizeof(int));
                DwmSetWindowAttribute(form.Handle, DwmwaTextColor, ref textColor, sizeof(int));
            }
            catch (DllNotFoundException) { }
            catch (EntryPointNotFoundException) { }
        }

        private static int ToColorRef(Color color)
        {
            return color.R | (color.G << 8) | (color.B << 16);
        }


        private void txtBrowser_Leave(object sender, EventArgs e)
        {
            if (txtBrowser.Text == String.Empty)
            {
                txtBrowser.Text = "Please select the directory for searching!!!";
                txtBrowser.ForeColor = Color.FromArgb(111, 137, 119);
            }
        }

        private void txtBrowser_Enter(object sender, EventArgs e)
        {
            if (txtBrowser.Text == "Please select the directory for searching!!!")
            {
                txtBrowser.Text = String.Empty;
                txtBrowser.ForeColor = Color.FromArgb(43, 74, 55);
            }
        }

        private void txtKeyWord_Enter(object sender, EventArgs e)
        {
            if (txtKeyWord.Text == "Enter the search keyword!!!")
            {
                txtKeyWord.Text = String.Empty;
                txtKeyWord.ForeColor = Color.FromArgb(43, 74, 55);
            }
        }

        private void txtKeyWord_Leave(object sender, EventArgs e)
        {
            if (txtKeyWord.Text == String.Empty)
            {
                txtKeyWord.Text = "Enter the search keyword!!!";
                txtKeyWord.ForeColor = Color.FromArgb(111, 137, 119);
            }
        }

        private void Form1_SizeChanged(object? sender, EventArgs e)
        {
            UpdateControlSizesAndLocations();
            UpdateProgressBarWidth();
            UpdateProgressBarPosition();
            UpdateProgressBarFont();
            BringToForntControl();
        }

        // ProcessBar
        private void InitializeProgressBarAndFileProcess()
        {
            emptyStatePanel.Visible = false;
            tableLayoutPanel.Visible = false;
            tableLayoutPanel.AutoScroll = false;
            resultsPagerHost.Visible = false;
            // Initialize ProgressBar
            progressBar1 = new ModernProgressBar
            {
                Minimum = 0,
                Maximum = 100,
                Step = 1,
                Visible = false,
                Height = 32,
                TrackColor = Color.FromArgb(216, 233, 221),
                ProgressColor = Color.FromArgb(103, 181, 130),
            };

            // Initialize Lable Progress Precent
            txtProgressPercent = new Label
            {
                TextAlign = ContentAlignment.MiddleCenter,
                BorderStyle = BorderStyle.None,
                Height = progressBar1.Height,
                Width = progressBar1.Width,
                BackColor = Color.FromArgb(244, 250, 245),
                ForeColor = Color.FromArgb(55, 112, 73),
            };
            // Initialize Lable Progress Detail
            txtProgressDetail = new Label
            {
                TextAlign = ContentAlignment.TopLeft,
                BorderStyle = BorderStyle.None,
                Height = progressBar1.Height,
                Width = progressBar1.Width,
                BackColor = Color.FromArgb(244, 250, 245),
                ForeColor = Color.FromArgb(83, 125, 96),
            };

            // Initialize Lable Result Path
            txtProgressFileHasKeyWord = new Label
            {
                TextAlign = ContentAlignment.TopRight,
                BorderStyle = BorderStyle.None,
                Height = progressBar1.Height,
                Width = progressBar1.Width,
                BackColor = Color.FromArgb(244, 250, 245),
                ForeColor = Color.FromArgb(83, 125, 96),
            };

            // Initialize Lable Current File
            txtProgressCurrentFile = new Label
            {
                TextAlign = ContentAlignment.TopLeft,
                BorderStyle = BorderStyle.None,
                Height = progressBar1.Height,
                Width = progressBar1.Width,
                BackColor = Color.FromArgb(244, 250, 245),
                ForeColor = Color.FromArgb(105, 133, 114),
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
            this.Controls.Add(txtProgressCurrentFile);

            // Bring ProgressBar to front
            progressBar1.BringToFront();
            txtProgressPercent.BringToFront();
            txtProgressDetail.BringToFront();
            txtProgressFileHasKeyWord.BringToFront();
            txtProgressCurrentFile.BringToFront();

            // Initialize FileProcess instance and subscribe to ProgressChanged event
            fileKeywordSearcher.ProgressChanged += FileProcessor_ProgressChanged;
        }

        private void UpdateProgressBarWidth()
        {
            if (progressBar1 != null && txtProgressPercent != null && txtProgressDetail != null && txtProgressFileHasKeyWord != null && txtProgressCurrentFile != null)
            {
                progressBar1.Width = Math.Min(ClientRectangle.Width - 120, 1400);
                progressBar1.Height = 32;

                txtProgressPercent.Width = progressBar1.Width;
                txtProgressPercent.Height = 38;

                txtProgressDetail.Width = progressBar1.Width / 2;
                txtProgressDetail.Height = txtProgressDetail.GetPreferredSize(new Size(txtProgressDetail.Width, int.MaxValue)).Height;

                txtProgressFileHasKeyWord.Width = progressBar1.Width / 2;
                txtProgressFileHasKeyWord.Height = 24;

                txtProgressCurrentFile.Width = progressBar1.Width;
                txtProgressCurrentFile.Height = 44;

            }
        }
        private void UpdateProgressBarPosition()
        {
            if (progressBar1 != null && txtProgressPercent != null && txtProgressDetail != null && txtProgressFileHasKeyWord != null && txtProgressCurrentFile != null)
            {
                int progressBarHeight = progressBar1.Height;
                int progressBarX = (ClientRectangle.Width - progressBar1.Width) / 2;
                int progressBarY = (ClientRectangle.Height - progressBarHeight) / 2;

                progressBar1.Location = new Point(progressBarX, progressBarY);
                txtProgressPercent.Location = new Point(progressBarX, progressBarY - txtProgressPercent.Height - 14);
                txtProgressDetail.Location = new Point(progressBarX, progressBarY + progressBar1.Height + 12);
                txtProgressFileHasKeyWord.Location = new Point(progressBarX + progressBar1.Width / 2, progressBarY + progressBar1.Height + 12);
                txtProgressCurrentFile.Location = new Point(progressBarX, txtProgressDetail.Location.Y + 28);
            }
        }

        private void UpdateProgressBarFont()
        {
            if (progressBar1 != null && txtProgressPercent != null && txtProgressDetail != null && txtProgressFileHasKeyWord != null && txtProgressCurrentFile != null)
            {
                txtProgressPercent.Font = new Font("Segoe UI", 16F, FontStyle.Bold);
                txtProgressDetail.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
                txtProgressFileHasKeyWord.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
                txtProgressCurrentFile.Font = new Font("Segoe UI", 8.5F, FontStyle.Regular);
            }
        }

        private void ControlsStatus(bool isEnable)
        {
            if (isEnable)
            {
                txtKeyWord.Enabled = true;
                txtBrowser.Enabled = true;
                btnBrowser.Enabled = true;
                labelWithCheckBoxList.Enabled = true;
                txtKeyWord.BackColor = Color.FromArgb(232, 242, 235);
                txtBrowser.BackColor = Color.FromArgb(232, 242, 235);
                btnBrowser.BackColor = Color.FromArgb(222, 240, 227);
                labelWithCheckBoxList.BackColor = Color.FromArgb(218, 238, 224);
                btnStartSearch.BackColor = Color.FromArgb(137, 201, 158);
                btnStartSearch.Text = "Search";
            }
            else
            {
                txtKeyWord.Enabled = false;
                txtBrowser.Enabled = false;
                btnBrowser.Enabled = false;
                labelWithCheckBoxList.Enabled = false;
                txtKeyWord.BackColor = Color.FromArgb(224, 233, 226);
                txtBrowser.BackColor = Color.FromArgb(224, 233, 226);
                btnBrowser.BackColor = Color.FromArgb(215, 226, 218);
                btnBrowser.FlatStyle = FlatStyle.Flat;
                btnBrowser.FlatAppearance.BorderSize = 0;
                labelWithCheckBoxList.BackColor = Color.FromArgb(215, 226, 218);
                btnStartSearch.BackColor = Color.FromArgb(238, 170, 160);
                btnStartSearch.Text = "Stop";
            }
        }

        private void ClearProgressBar()
        {
            if (fileKeywordSearcher != null)
                fileKeywordSearcher.ProgressChanged -= FileProcessor_ProgressChanged;
            if (progressBar1 != null)
            {
                progressBar1.Visible = false;
                this.Controls.Remove(progressBar1);
                progressBar1 = null;
            }

            if (txtProgressPercent != null)
            {
                txtProgressPercent.Visible = false;
                this.Controls.Remove(txtProgressPercent);
                txtProgressPercent = null;
            }

            if (txtProgressDetail != null)
            {
                txtProgressDetail.Visible = false;
                this.Controls.Remove(txtProgressDetail);
                txtProgressDetail = null;
            }

            if (txtProgressFileHasKeyWord != null)
            {
                txtProgressFileHasKeyWord.Visible = false;
                this.Controls.Remove(txtProgressFileHasKeyWord);
                txtProgressFileHasKeyWord = null;
            }

            if (txtProgressCurrentFile != null)
            {
                txtProgressCurrentFile.Visible = false;
                this.Controls.Remove(txtProgressCurrentFile);
                txtProgressCurrentFile = null;
            }

            InitializeTableLayoutResult();
            ControlsStatus(true);
        }

        private void BringToForntControl()
        {
            btnBrowser.BringToFront();
            txtBrowser.BringToFront();
            txtKeyWord.BringToFront();
            btnBrowser.BringToFront();
        }
    }
}
