namespace FileKeywordSearcher
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && components != null) components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new(typeof(Form1));
            rootLayout = new TableLayoutPanel();
            resultsCard = new RoundedPanel();
            resultsContentLayout = new TableLayoutPanel();
            tableLayoutPanel = new TableLayoutPanel();
            resultsPagerHost = new Panel();
            emptyStatePanel = new TableLayoutPanel();
            lblEmptyIcon = new Label();
            lblEmptyTitle = new Label();
            lblEmptyText = new Label();
            searchCard = new RoundedPanel();
            searchLayout = new TableLayoutPanel();
            lblKeyword = new Label();
            lblFolder = new Label();
            txtKeyWord = new TextBox();
            txtBrowser = new TextBox();
            labelWithCheckBoxList = new LabelWithCheckBoxList();
            btnBrowser = new ModernButton();
            btnStartSearch = new ModernButton();
            rootLayout.SuspendLayout();
            resultsCard.SuspendLayout();
            resultsContentLayout.SuspendLayout();
            emptyStatePanel.SuspendLayout();
            searchCard.SuspendLayout();
            searchLayout.SuspendLayout();
            SuspendLayout();

            rootLayout.BackColor = Color.Transparent;
            rootLayout.ColumnCount = 1;
            rootLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            rootLayout.Controls.Add(resultsCard, 0, 0);
            rootLayout.Controls.Add(searchCard, 0, 1);
            rootLayout.Dock = DockStyle.Fill;
            rootLayout.Padding = new Padding(28, 20, 28, 24);
            rootLayout.RowCount = 2;
            rootLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 148F));

            resultsCard.BackColor = Color.FromArgb(244, 250, 245);
            resultsCard.BorderColor = Color.FromArgb(218, 233, 221);
            resultsCard.HighlightColor = Color.White;
            resultsCard.ShadowColor = Color.FromArgb(190, 212, 196);
            resultsCard.ShowLeafPattern = true;
            resultsCard.CornerRadius = 22;
            resultsCard.Controls.Add(resultsContentLayout);
            resultsCard.Controls.Add(emptyStatePanel);
            resultsCard.Dock = DockStyle.Fill;
            resultsCard.Margin = new Padding(0, 0, 0, 16);
            resultsCard.Padding = new Padding(14);

            resultsContentLayout.BackColor = Color.Transparent;
            resultsContentLayout.ColumnCount = 1;
            resultsContentLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            resultsContentLayout.Controls.Add(tableLayoutPanel, 0, 0);
            resultsContentLayout.Controls.Add(resultsPagerHost, 0, 1);
            resultsContentLayout.Dock = DockStyle.Fill;
            resultsContentLayout.RowCount = 2;
            resultsContentLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            resultsContentLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 58F));

            tableLayoutPanel.AutoScroll = true;
            tableLayoutPanel.BackColor = Color.Transparent;
            tableLayoutPanel.ColumnCount = 1;
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel.Dock = DockStyle.Fill;
            tableLayoutPanel.Visible = false;

            resultsPagerHost.BackColor = Color.FromArgb(244, 250, 245);
            resultsPagerHost.Dock = DockStyle.Fill;
            resultsPagerHost.Visible = false;

            emptyStatePanel.BackColor = Color.Transparent;
            emptyStatePanel.ColumnCount = 1;
            emptyStatePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            emptyStatePanel.Controls.Add(lblEmptyIcon, 0, 1);
            emptyStatePanel.Controls.Add(lblEmptyTitle, 0, 2);
            emptyStatePanel.Controls.Add(lblEmptyText, 0, 3);
            emptyStatePanel.Dock = DockStyle.Fill;
            emptyStatePanel.RowCount = 5;
            emptyStatePanel.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            emptyStatePanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 58F));
            emptyStatePanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));
            emptyStatePanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            emptyStatePanel.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            lblEmptyIcon.Dock = DockStyle.Fill;
            lblEmptyIcon.Font = new Font("Segoe UI Symbol", 28F);
            lblEmptyIcon.ForeColor = Color.FromArgb(91, 169, 119);
            lblEmptyIcon.Text = "⌕";
            lblEmptyIcon.TextAlign = ContentAlignment.MiddleCenter;
            lblEmptyTitle.Dock = DockStyle.Fill;
            lblEmptyTitle.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            lblEmptyTitle.ForeColor = Color.FromArgb(43, 74, 55);
            lblEmptyTitle.Text = "Ready when you are";
            lblEmptyTitle.TextAlign = ContentAlignment.MiddleCenter;
            lblEmptyText.Dock = DockStyle.Fill;
            lblEmptyText.Font = new Font("Segoe UI", 9.5F);
            lblEmptyText.ForeColor = Color.FromArgb(105, 133, 114);
            lblEmptyText.Text = "Choose a folder, enter a keyword, then start searching.";
            lblEmptyText.TextAlign = ContentAlignment.MiddleCenter;

            searchCard.BackColor = Color.FromArgb(240, 248, 242);
            searchCard.BorderColor = Color.FromArgb(213, 230, 217);
            searchCard.HighlightColor = Color.White;
            searchCard.ShadowColor = Color.FromArgb(185, 208, 191);
            searchCard.ShowLeafPattern = true;
            searchCard.CornerRadius = 22;
            searchCard.Controls.Add(searchLayout);
            searchCard.Dock = DockStyle.Fill;
            searchCard.Padding = new Padding(18, 12, 18, 12);

            searchLayout.BackColor = Color.Transparent;
            searchLayout.ColumnCount = 3;
            searchLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            searchLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 122F));
            searchLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 122F));
            searchLayout.Controls.Add(lblKeyword, 0, 0);
            searchLayout.Controls.Add(lblFolder, 0, 2);
            searchLayout.Controls.Add(txtKeyWord, 0, 1);
            searchLayout.Controls.Add(labelWithCheckBoxList, 1, 1);
            searchLayout.Controls.Add(btnStartSearch, 2, 1);
            searchLayout.Controls.Add(txtBrowser, 0, 3);
            searchLayout.Controls.Add(btnBrowser, 1, 3);
            searchLayout.Dock = DockStyle.Fill;
            searchLayout.RowCount = 4;
            searchLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 21F));
            searchLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 37F));
            searchLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 21F));
            searchLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 37F));

            lblKeyword.AutoSize = true; lblKeyword.ForeColor = Color.FromArgb(83, 125, 96); lblKeyword.Text = "KEYWORD"; lblKeyword.Font = new Font("Segoe UI", 8F, FontStyle.Bold);
            lblFolder.AutoSize = true; lblFolder.ForeColor = Color.FromArgb(83, 125, 96); lblFolder.Text = "SEARCH LOCATION"; lblFolder.Font = new Font("Segoe UI", 8F, FontStyle.Bold);

            txtKeyWord.BackColor = Color.FromArgb(232, 242, 235); txtKeyWord.BorderStyle = BorderStyle.FixedSingle; txtKeyWord.Dock = DockStyle.Fill; txtKeyWord.Font = new Font("Segoe UI", 10F); txtKeyWord.ForeColor = Color.FromArgb(111, 137, 119); txtKeyWord.Margin = new Padding(0, 0, 10, 4); txtKeyWord.Padding = new Padding(8); txtKeyWord.Text = "Enter the search keyword!!!"; txtKeyWord.Enter += txtKeyWord_Enter; txtKeyWord.Leave += txtKeyWord_Leave;
            txtBrowser.BackColor = Color.FromArgb(232, 242, 235); txtBrowser.BorderStyle = BorderStyle.FixedSingle; txtBrowser.Dock = DockStyle.Fill; txtBrowser.Font = new Font("Segoe UI", 10F); txtBrowser.ForeColor = Color.FromArgb(111, 137, 119); txtBrowser.Margin = new Padding(0, 0, 10, 2); txtBrowser.Text = "Please select the directory for searching!!!"; txtBrowser.Enter += txtBrowser_Enter; txtBrowser.Leave += txtBrowser_Leave;

            labelWithCheckBoxList.BackColor = Color.FromArgb(218, 238, 224); labelWithCheckBoxList.Dock = DockStyle.Fill; labelWithCheckBoxList.Font = new Font("Segoe UI", 9F, FontStyle.Bold); labelWithCheckBoxList.ForeColor = Color.FromArgb(55, 100, 69); labelWithCheckBoxList.Margin = new Padding(0, 0, 10, 4); labelWithCheckBoxList.Text = "All"; labelWithCheckBoxList.TextAlign = ContentAlignment.MiddleCenter;
            btnBrowser.BackColor = Color.FromArgb(222, 240, 227); btnBrowser.BorderColor = Color.FromArgb(194, 220, 202); btnBrowser.CornerRadius = 10; btnBrowser.Dock = DockStyle.Fill; btnBrowser.FlatStyle = FlatStyle.Flat; btnBrowser.Font = new Font("Segoe UI", 9F, FontStyle.Bold); btnBrowser.ForeColor = Color.FromArgb(55, 100, 69); btnBrowser.Margin = new Padding(0, 0, 10, 2); btnBrowser.Text = "Browse…"; btnBrowser.Click += btnBrowser_Click;
            btnStartSearch.BackColor = Color.FromArgb(137, 201, 158); btnStartSearch.BorderColor = Color.FromArgb(116, 185, 139); btnStartSearch.CornerRadius = 11; btnStartSearch.Dock = DockStyle.Fill; btnStartSearch.FlatStyle = FlatStyle.Flat; btnStartSearch.Font = new Font("Segoe UI", 10F, FontStyle.Bold); btnStartSearch.ForeColor = Color.FromArgb(28, 73, 43); btnStartSearch.Margin = new Padding(0, 0, 0, 4); searchLayout.SetRowSpan(btnStartSearch, 3); btnStartSearch.Text = "Search"; btnStartSearch.Click += btnStartSearch_Click_1;

            AutoScaleDimensions = new SizeF(8F, 20F); AutoScaleMode = AutoScaleMode.Font; BackColor = Color.FromArgb(232, 243, 235); ClientSize = new Size(980, 680); Controls.Add(rootLayout); DoubleBuffered = true; Font = new Font("Segoe UI", 9F); ForeColor = Color.FromArgb(43, 74, 55); Icon = (Icon)resources.GetObject("$this.Icon"); MinimumSize = new Size(760, 560); Name = "Form1"; Text = "File Search — Local keyword search";
            rootLayout.ResumeLayout(false); resultsCard.ResumeLayout(false); resultsContentLayout.ResumeLayout(false); emptyStatePanel.ResumeLayout(false); searchCard.ResumeLayout(false); searchLayout.ResumeLayout(false); searchLayout.PerformLayout(); ResumeLayout(false);
        }

        private TableLayoutPanel rootLayout, searchLayout, tableLayoutPanel, emptyStatePanel, resultsContentLayout;
        private Panel resultsPagerHost;
        private RoundedPanel resultsCard, searchCard;
        private Label lblEmptyIcon, lblEmptyTitle, lblEmptyText, lblKeyword, lblFolder;
        private TextBox txtKeyWord, txtBrowser;
        private ModernButton btnBrowser, btnStartSearch;
        private LabelWithCheckBoxList labelWithCheckBoxList;
    }
}
