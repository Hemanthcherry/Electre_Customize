using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Electre_Customize_DotNet.Contracts
{
    internal interface ISelectionFunctions
    {
         void SelectCheckItems(CheckBox checkbox, CheckedListBox chkListbox, HashSet<string> checkedItem);

         void SelectCount(Label lblSelect, HashSet<string> selectedCount);

         void TextFilter(TextBox txtFilter, HashSet<string> chekedItems, CheckedListBox chkList, List<string> originalList);

         void DefualtFilterName(TextBox textFilter);

        void ListItemChecked(ItemCheckEventArgs e, HashSet<string> checkedItem, CheckedListBox chkList, Label lblCount);
        void ListItemChecked(ItemCheckEventArgs e, HashSet<string> checkedItem, CheckedListBox chkList, Label lblCount, CheckBox selectAllCheckbox);

        void ListItemCheckedPanelDwg(ItemCheckEventArgs e, HashSet<string> checkedItem, CheckedListBox chkList, CheckBox selectAllCheckbox);

        void UpdateCheckedListBox(List<string> filteredItems, HashSet<string> checkedItems, CheckedListBox chkListBox);

        void TextClear(TextBox txtFilter);
       // void UpdateListBox(List<string> arrListOfLOOM, HashSet<string> checkedLoomItems, ListBox listBox);
    }
}
