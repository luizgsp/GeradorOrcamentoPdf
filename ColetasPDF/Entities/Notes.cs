namespace ColetasPDF.Entities
{
    class Notes
    {
        public string Name { get; set; }
        public string Value { get; set; }

        public Notes(string name, string value)
        {
            Name = name;
            Value = value;
        }
    }
}
