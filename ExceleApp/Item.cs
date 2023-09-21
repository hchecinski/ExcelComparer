namespace ExceleApp
{
    public class Item
    {
        public Item()
        {

        }
        public Item(string value)
        {
            Value = value;
        }

        public string Value { get; set; }
        public bool IsNewValue { get; set; }
        public override string ToString()
        {
            return $"{Value}, {IsNewValue}";
        }
    }
}
