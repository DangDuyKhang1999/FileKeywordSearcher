using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;

namespace FileKeywordSearcher
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btnBrowser = new Button();
            txtBrowser = new TextBox();
            tableLayoutPanel = new TableLayoutPanel();
            btnStartSearch = new Button();
            SuspendLayout();
            // 
            // btnBrowser
            // 
            btnBrowser.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnBrowser.BackColor = Color.FromArgb(137, 190, 179);
            btnBrowser.FlatAppearance.BorderColor = Color.FromArgb(137, 190, 179);
            btnBrowser.FlatStyle = FlatStyle.Flat;
            btnBrowser.ForeColor = Color.Black;
            btnBrowser.Location = new Point(694, 3);
            btnBrowser.Name = "btnBrowser";
            btnBrowser.Size = new Size(103, 40);
            btnBrowser.TabIndex = 1;
            btnBrowser.Text = "Browser";
            btnBrowser.UseVisualStyleBackColor = false;
            btnBrowser.Click += BtnBrowser_Click;
            // 
            // txtBrowser
            // 
            txtBrowser.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left;
            txtBrowser.BackColor = Color.FromArgb(190, 217, 217);
            txtBrowser.BorderStyle = BorderStyle.None;
            txtBrowser.ForeColor = Color.Red;
            txtBrowser.Location = new Point(3, 3);
            txtBrowser.Multiline = true;
            txtBrowser.Name = "txtBrowser";
            txtBrowser.Size = new Size(689, 40);
            txtBrowser.TabIndex = 2;
            txtBrowser.Text = "Please select the directory for inspection!!!";
            txtBrowser.Enter += txtBrowser_Enter;
            txtBrowser.Leave += txtBrowser_Leave;
            // 
            // tableLayoutPanel
            // 
            tableLayoutPanel.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left;
            tableLayoutPanel.AutoScroll = true;
            tableLayoutPanel.BackColor = Color.FromArgb(190, 217, 217);
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20F));
            tableLayoutPanel.ForeColor = Color.FromArgb(161, 204, 209);
            tableLayoutPanel.Location = new Point(3, 45);
            tableLayoutPanel.Name = "tableLayoutPanel";
            tableLayoutPanel.Size = new Size(689, 400);
            tableLayoutPanel.TabIndex = 6;
            // 
            // btnStartSearch
            // 
            btnStartSearch.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnStartSearch.BackColor = Color.FromArgb(137, 190, 179);
            btnStartSearch.FlatAppearance.BorderColor = Color.FromArgb(137, 190, 179);
            btnStartSearch.FlatStyle = FlatStyle.Flat;
            btnStartSearch.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            btnStartSearch.ForeColor = Color.Black;
            btnStartSearch.Location = new Point(694, 385);
            btnStartSearch.Name = "btnStartSearch";
            btnStartSearch.Size = new Size(103, 60);
            btnStartSearch.TabIndex = 3;
            btnStartSearch.Text = "Search";
            btnStartSearch.UseVisualStyleBackColor = false;
            btnStartSearch.Click += btnStartSearch_Click_1;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.FromArgb(238, 245, 255);
            ClientSize = new Size(800, 449);
            Controls.Add(btnStartSearch);
            Controls.Add(tableLayoutPanel);
            Controls.Add(txtBrowser);
            Controls.Add(btnBrowser);
            Name = "Form1";
            Text = "File Keyword Searcher";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnBrowser;
        private TextBox txtBrowser;
        //private RichTextBox rtxtResult;
        private TableLayoutPanel tableLayoutPanel;
        private Button btnStartSearch;
    }
}