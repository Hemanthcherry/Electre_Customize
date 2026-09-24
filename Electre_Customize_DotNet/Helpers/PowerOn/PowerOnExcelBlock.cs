namespace Electre_Customize_DotNet.Helpers.PowerOn
{
    internal sealed class PowerOnExcelBlock
    {
        public int StartRow;
        public int TotalRows;
        public object[,] Rows = new object[0, 0];
        public bool MergePanelScb;
        public bool MergePanelTcb;
    }
}
