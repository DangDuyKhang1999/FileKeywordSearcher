﻿using System.Windows.Forms;
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            btnBrowser = new Button();
            txtBrowser = new TextBox();
            tableLayoutPanel = new TableLayoutPanel();
            btnStartSearch = new Button();
            txtKeyWord = new TextBox();
            labelWithCheckBoxList = new LabelWithCheckBoxList();
            SuspendLayout();
            // 
            // btnBrowser
            // 
            btnBrowser.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnBrowser.BackColor = Color.FromArgb(137, 190, 179);
            btnBrowser.FlatAppearance.BorderColor = Color.FromArgb(137, 190, 179);
            btnBrowser.FlatStyle = FlatStyle.Flat;
            btnBrowser.ForeColor = Color.Black;
            btnBrowser.Location = new Point(589, 416);
            btnBrowser.Name = "btnBrowser";
            btnBrowser.Size = new Size(103, 29);
            btnBrowser.TabIndex = 2;
            btnBrowser.Text = "Browser";
            btnBrowser.UseVisualStyleBackColor = false;
            btnBrowser.Click += btnBrowser_Click;
            // 
            // txtBrowser
            // 
            txtBrowser.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            txtBrowser.BackColor = Color.FromArgb(190, 217, 217);
            txtBrowser.BorderStyle = BorderStyle.None;
            txtBrowser.ForeColor = Color.Red;
            txtBrowser.Location = new Point(3, 416);
            txtBrowser.Multiline = true;
            txtBrowser.Name = "txtBrowser";
            txtBrowser.Size = new Size(584, 29);
            txtBrowser.TabIndex = 4;
            txtBrowser.Text = "Please select the directory for searching!!!";
            txtBrowser.Enter += txtBrowser_Enter;
            txtBrowser.Leave += txtBrowser_Leave;
            // 
            // tableLayoutPanel
            // 
            tableLayoutPanel.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            tableLayoutPanel.AutoScroll = true;
            tableLayoutPanel.BackColor = Color.FromArgb(190, 217, 217);
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20F));
            tableLayoutPanel.ForeColor = Color.FromArgb(161, 204, 209);
            tableLayoutPanel.Location = new Point(3, 3);
            tableLayoutPanel.Name = "tableLayoutPanel";
            tableLayoutPanel.Size = new Size(794, 380);
            tableLayoutPanel.TabIndex = 1;
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
            // txtKeyWord
            // 
            txtKeyWord.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            txtKeyWord.BackColor = Color.FromArgb(190, 217, 217);
            txtKeyWord.BorderStyle = BorderStyle.None;
            txtKeyWord.ForeColor = Color.Red;
            txtKeyWord.Location = new Point(3, 385);
            txtKeyWord.Multiline = true;
            txtKeyWord.Name = "txtKeyWord";
            txtKeyWord.Size = new Size(584, 29);
            txtKeyWord.TabIndex = 1;
            txtKeyWord.Text = "Enter the search keyword!!!";
            txtKeyWord.Enter += txtKeyWord_Enter;
            txtKeyWord.Leave += txtKeyWord_Leave;
            // 
            // labelWithCheckBoxList
            // 
            labelWithCheckBoxList.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            labelWithCheckBoxList.BackColor = Color.FromArgb(137, 190, 179);
            labelWithCheckBoxList.Location = new Point(589, 385);
            labelWithCheckBoxList.Name = "labelWithCheckBoxList";
            labelWithCheckBoxList.Size = new Size(103, 29);
            labelWithCheckBoxList.TabIndex = 5;
            labelWithCheckBoxList.Text = "All";
            labelWithCheckBoxList.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.FromArgb(238, 245, 255);
            ClientSize = new Size(800, 449);
            Controls.Add(txtKeyWord);
            Controls.Add(btnStartSearch);
            Controls.Add(tableLayoutPanel);
            Controls.Add(txtBrowser);
            Controls.Add(btnBrowser);
            Controls.Add(labelWithCheckBoxList);
            Icon = (Icon)resources.GetObject("$this.Icon");
            MinimumSize = new Size(550, 300);
            Name = "Form1";
            Text = "File Keyword Searcher";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TextBox txtKeyWord;
        private TextBox txtBrowser;
        private TableLayoutPanel tableLayoutPanel;
        private Button btnBrowser;
        private Button btnStartSearch;
        private LabelWithCheckBoxList labelWithCheckBoxList;
    }
}