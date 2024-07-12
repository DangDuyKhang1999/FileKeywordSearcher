using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using static FileKeywordSearcher.Form1;

namespace FileKeywordSearcher
{
    public class LabelWithCheckBoxList : Label
    {
        private CheckedListBox _checkedListBox;
        private Form _popupForm;
        private bool _isDroppedDown;
        private bool initCheck = false;
        public List<object> SelectedItems { get; } = new List<object>();

        public LabelWithCheckBoxList()
        {
            _checkedListBox = new CheckedListBox
            {
                CheckOnClick = true
            };
            _checkedListBox.ItemCheck += CheckedListBox_ItemCheck;

            _popupForm = new Form
            {
                FormBorderStyle = FormBorderStyle.None,
                StartPosition = FormStartPosition.Manual,
                ShowInTaskbar = false,
                AutoSizeMode = AutoSizeMode.GrowOnly
            };
            _popupForm.Controls.Add(_checkedListBox);

            this.Click += LabelWithCheckBoxList_Click;
        }

        private void LabelWithCheckBoxList_Click(object sender, EventArgs e)
        {
            if (initCheck)
            {
                ShowPopup();
            }
            else
            {
                // :D
                _popupForm.Show();
                _popupForm.Hide();
                initCheck = true;
                ShowPopup();
            }
        }

        private void ShowPopup()
        {
            if (!_isDroppedDown)
            {
                _checkedListBox.Items.Clear();
                foreach (var item in Enum.GetValues(typeof(eTargetExtension)))
                {
                    _checkedListBox.Items.Add(item, this.SelectedItems.Contains(item));
                }

                // Calculate the height needed to display all items without scroll
                int totalHeight = _checkedListBox.ItemHeight * _checkedListBox.Items.Count;
                totalHeight = Math.Min(totalHeight, Screen.PrimaryScreen.WorkingArea.Height / 2); // Limit height if too tall

                // Set checkedListBox size
                _checkedListBox.Size = new Size(103, totalHeight);

                // Set popupForm size
                _popupForm.ClientSize = new Size(103, _checkedListBox.Height);

                // Set maximum width for the popup form
                _popupForm.MaximumSize = new Size(103, _checkedListBox.Height);

                var point = this.PointToScreen(new Point(0, this.Height));
                _popupForm.Location = point;

                _popupForm.Show();
                _isDroppedDown = true;
            }
            else
            {
                _popupForm.Hide();
                _isDroppedDown = false;
            }
        }

        private void CheckedListBox_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (e.NewValue == CheckState.Checked)
            {
                this.SelectedItems.Add(_checkedListBox.Items[e.Index]);
            }
            else
            {
                this.SelectedItems.Remove(_checkedListBox.Items[e.Index]);
            }

            UpdateLabelText();
        }

        private void UpdateLabelText()
        {
            if (this.SelectedItems.Count > 2)
            {
                // Display only the first item followed by ",..."
                this.Text = this.SelectedItems.FirstOrDefault()?.ToString() + ",...";
            }
            else
            {
                // Display all selected items separated by ","
                this.Text = string.Join(", ", this.SelectedItems);
            }
        }
    }
}
