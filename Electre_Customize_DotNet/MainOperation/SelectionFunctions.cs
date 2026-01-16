using Electre_Customize_DotNet.Contracts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Electre_Customize_DotNet.MainOperation
{
    internal class SelectionFunctions: ISelectionFunctions
    {
        private static SelectionFunctions selctionInstance;

        private static readonly object selctionInstanceLock = new object(); 

        private SelectionFunctions() { }

        public static SelectionFunctions SelectionFuntionsInstance
        {
            get
            {
                lock (selctionInstanceLock)
                {
                    if(selctionInstance == null)
                    {
                        selctionInstance = new SelectionFunctions();
                    }

                    return selctionInstance;
                }
            }
        }

        public void SelectCheckItems(CheckBox checkbox, CheckedListBox chkListbox, HashSet<string> checkedItem)
        {
            
            if (checkbox.Checked == true)
            {
                bool isChecked = checkbox.Checked;

                for (int i = 0; i < chkListbox.Items.Count; i++)
                {
                    chkListbox.SetItemChecked(i, isChecked);
                    string item = chkListbox.Items[i].ToString();
                    if (isChecked)
                    {
                        checkedItem.Add(item);
                    }
                    else
                    {
                        checkedItem.Remove(item);
                    }
                }
            }

            else
            {
                for (int i = 0; i < chkListbox.Items.Count; i++)
                {
                    chkListbox.SetItemChecked(i, false);
                }
            }

        }
       //public  int TotalCount(int count)
       //{

       //     return count;
        
       //}

        public void SelectCount(Label lblSelect, HashSet<string> selectedCount)
        {
            string searchCount = lblSelect.Text;
            lblSelect.Text = "Selected Items Count: " + selectedCount.Count;
        }

        public void TextFilter(TextBox txtFilter, HashSet<string> chekedItems, CheckedListBox chkList, List<string> originalList)
        {
            if (txtFilter.Text.ToUpper() == "SEARCH FILTER")
            {
                return;
            }

            else
            {
                txtFilter.Font =new Font("Segoe UI", 10F);
                txtFilter.ForeColor = Color.Black;
                string filter = txtFilter.Text.ToLower();
                var filteredItems = originalList.Where(item => item.ToLower().Contains(filter)).ToList();
                UpdateCheckedListBox(filteredItems, chekedItems, chkList);
            }
        }

        public void DefualtFilterName(TextBox textFilter)
        {
            if (textFilter.Text == "" || textFilter.Text == null)
            {
                textFilter.Text = "Search Filter";
                textFilter.Font = new Font("Segoe UI", 12F, FontStyle.Italic);
                textFilter.ForeColor = Color.Silver;
            }
        }
        // Commented on June 9th before implementing SelectAll intermediate functionality
        public void ListItemChecked(ItemCheckEventArgs e, HashSet<string> checkedItem, CheckedListBox chkList, Label lblCount)//add and remove the items to track.
        {
            string item = chkList.Items[e.Index].ToString();
            if (e.NewValue == CheckState.Checked)
            {
                checkedItem.Add(item);
                SelectCount(lblCount, checkedItem);
            }
            else
            {
                checkedItem.Remove(item);
                SelectCount(lblCount, checkedItem);
            }

        }

        public void ListItemChecked(ItemCheckEventArgs e, HashSet<string> checkedItem, CheckedListBox chkList, Label lblCount, CheckBox selectAllCheckbox)
        {
            /*if (SuppressListItemEvents)
                return;*/

            string item = chkList.Items[e.Index].ToString();

            if (e.NewValue == CheckState.Checked)
                checkedItem.Add(item);
            else
                checkedItem.Remove(item);

            SelectCount(lblCount, checkedItem);

            // Safely update Select All checkbox
            if (checkedItem.Count == chkList.Items.Count)
                selectAllCheckbox.CheckState = CheckState.Checked;
            else if (checkedItem.Count == 0)
                selectAllCheckbox.CheckState = CheckState.Unchecked;
            else
                selectAllCheckbox.CheckState = CheckState.Indeterminate;

            chkList.ClearSelected(); // removes blue highlight

        }

        public void ListItemCheckedPanelDwg(ItemCheckEventArgs e, HashSet<string> checkedItem, CheckedListBox chkList, CheckBox selectAllCheckbox)
        {
            string item = chkList.Items[e.Index].ToString();

            if (e.NewValue == CheckState.Checked)
                checkedItem.Add(item);
            else
                checkedItem.Remove(item);

            // Safely update Select All checkbox
            if (checkedItem.Count == chkList.Items.Count)
                selectAllCheckbox.CheckState = CheckState.Checked;
            else if (checkedItem.Count == 0)
                selectAllCheckbox.CheckState = CheckState.Unchecked;
            else
                selectAllCheckbox.CheckState = CheckState.Indeterminate;

            chkList.ClearSelected(); // removes blue highlight
        }

        // commented on July 10th before implementing flickering issue.
        /* public void ListItemChecked(ItemCheckEventArgs e, HashSet<string> checkedItem, CheckedListBox chkList, Label lblCount, CheckBox selectAllCheckbox) // NEW parameter
         {
             if (SuppressListItemEvents)
                 return;
             string item = chkList.Items[e.Index].ToString();
             if (e.NewValue == CheckState.Checked)
             {
                 checkedItem.Add(item);
             }
             else
             {
                 checkedItem.Remove(item);
             }

             // Update label
             SelectCount(lblCount, checkedItem);

             // Delay the evaluation till after UI updates
             chkList.BeginInvoke((MethodInvoker)(() =>
             {
                 if (SuppressListItemEvents) return;
                 int checkedCount = chkList.CheckedItems.Count;
                 if (chkList.GetItemCheckState(e.Index) != CheckState.Checked && e.NewValue == CheckState.Checked)
                 {
                     checkedCount++; // item is going to be checked
                 }
                 else if (chkList.GetItemCheckState(e.Index) == CheckState.Checked && e.NewValue == CheckState.Unchecked)
                 {
                     checkedCount--; // item is going to be unchecked
                 }

                 if (checkedCount == chkList.Items.Count)
                     selectAllCheckbox.CheckState = CheckState.Checked;
                 else if (checkedCount == 0)
                     selectAllCheckbox.CheckState = CheckState.Unchecked;
                 else
                     selectAllCheckbox.CheckState = CheckState.Indeterminate;
             }));
         }*/


        public void UpdateCheckedListBox(List<string> filteredItems, HashSet<string> checkedItems, CheckedListBox chkListBox)
        {
            if (checkedItems == null)
                checkedItems = new HashSet<string>();

            chkListBox.Items.Clear();
            foreach (var item in filteredItems)
            {
                chkListBox.Items.Add(item, checkedItems.Contains(item));
            }
        }

        public void TextClear(TextBox txtFilter)
        {
            if (txtFilter.Text.ToUpper() == "SEARCH FILTER")
            {
                txtFilter.Text = string.Empty;
            }
        }
    }
}
