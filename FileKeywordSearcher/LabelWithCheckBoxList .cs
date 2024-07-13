using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace FileKeywordSearcher
{
    public class LabelWithCheckBoxList : Label
    {
        private CheckedListBox _checkedListBox;
        private Form _popupForm;
        private bool _isDroppedDown;
        private bool initCheck = false;
        public HashSet<eTargetExtension> m_SelectedItems { get; } = new HashSet<eTargetExtension>();

        // Property to store the parent form
        public Form ParentForm { get; private set; }

        public LabelWithCheckBoxList()
        {
            _checkedListBox = new CheckedListBox
            {
                CheckOnClick = true,
                BackColor = Color.FromArgb(190, 217, 217),
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
            _popupForm.Deactivate += PopupForm_Deactivate;
            this.Click += LabelWithCheckBoxList_Click;

            // Subscribe to parent form's Load event to initialize ParentForm property
            this.ParentChanged += (sender, e) =>
            {
                ParentForm = this.FindForm();
            };
        }


        private void PopupForm_Deactivate(object sender, EventArgs e)
        {
            HidePopup();
        }

        private void LabelWithCheckBoxList_Click(object sender, EventArgs e)
        {
            if (initCheck)
            {
                ShowPopup();
            }
            else
            {
                _popupForm.Show();
                _popupForm.Hide();
                initCheck = true;
                ShowPopup();
                if (_checkedListBox.Items.Count > 0)
                {
                    _checkedListBox.SetItemChecked(0, true); // Last item
                }
            }

            // Example: Accessing parent form's height
            if (ParentForm != null)
            {
                int formHeight = ParentForm.Height;
                // Now you can use 'formHeight' as needed
                // For example:
                Console.WriteLine($"Parent form height: {formHeight}");
            }
        }

        private void ShowPopup()
        {
            if (!_isDroppedDown)
            {
                _checkedListBox.Items.Clear();
                foreach (var item in Enum.GetValues(typeof(eTargetExtension)))
                {
                    _checkedListBox.Items.Add(item, this.m_SelectedItems.Contains((eTargetExtension)item));
                }

                int totalHeight = _checkedListBox.ItemHeight * (_checkedListBox.Items.Count + 1);
                totalHeight = Math.Min(totalHeight, Screen.PrimaryScreen.WorkingArea.Height / 2);

                _checkedListBox.Size = new Size(103, totalHeight);
                _popupForm.ClientSize = new Size(103, _checkedListBox.Height);
                _popupForm.MaximumSize = new Size(103, _checkedListBox.Height);

                var point = this.PointToScreen(new Point(0, this.Height));
                if (IsFullScreen() || IsInLowerHalfOfScreen())
                {
                    point = this.PointToScreen(new Point(0, -_popupForm.Height));
                }
                _popupForm.Location = point;

                _popupForm.Show();
                _isDroppedDown = true;
            }
            else
            {
                HidePopup();
            }
        }

        private void HidePopup()
        {
            if (_isDroppedDown)
            {
                _popupForm.Hide();
                _isDroppedDown = false;
            }
        }

        private void CheckedListBox_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            var selectedItem = (eTargetExtension)_checkedListBox.Items[e.Index];

            if (e.NewValue == CheckState.Checked)
            {
                this.m_SelectedItems.Add(selectedItem);
            }
            else
            {
                this.m_SelectedItems.Remove(selectedItem);
            }

            if (m_SelectedItems.Count > 0)
            {
                UpdateLabelText();
            }
            else
            {
                this.Text = "All";
                m_SelectedItems.Clear();
            }
        }

        private void UpdateLabelText()
        {
            if (this.m_SelectedItems.Count > 2)
            {
                var firstItem = this.m_SelectedItems.FirstOrDefault();
                this.Text = firstItem + ",...";
            }
            else if (this.m_SelectedItems.Count == 0)
            {
                this.Text = "All";
            }
            else
            {
                this.Text = string.Join(", ", this.m_SelectedItems);
            }
        }

        public bool IsFullScreen()
        {
            if (ParentForm != null)
            {
                int screenHeight = Screen.PrimaryScreen.Bounds.Height;
                return ParentForm.Height >= screenHeight - 100;
            }
            return false;
        }
        public bool IsInLowerHalfOfScreen()
        {
            if (ParentForm != null)
            {
                Rectangle screenRectangle = Screen.FromControl(ParentForm).Bounds;
                int screenHeight = screenRectangle.Height;

                int formBottom = ParentForm.Location.Y + ParentForm.Height/2;

                return (formBottom >= screenHeight / 2);
            }
            return false;
        }
    }
}
