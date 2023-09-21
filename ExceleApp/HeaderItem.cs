namespace ExceleApp
{
    public class HeaderItem
    {
        public HeaderItem()
        {

        }
        public HeaderItem(int columnIndex, string name)
        {
            ColumnIndex = columnIndex;
            Value = name;
        }

        private int _columnIndex = -1;
        public int ColumnIndex { get => _columnIndex; set => _columnIndex = value - 1; }
        public string Value { get; set; }
        public bool IsNew { get; set; }
        public override string ToString()
        {
            return $"{Value}";
        }
    }
}
