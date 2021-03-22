namespace ColetasPDF.Entities
{
    class Item
    {
        public int ItemSequel { get; set; }
        public string ItemCode { get; set; }
        public string Description { get; set; }
        public double Quantity { get; set; }
        public string TaxClassification { get; set; }
        public string IcmsTax { get; set; }
        public double UnitPrice { get; set; }
        public string IpiTax { get; set; }
        public string DeliverTime { get; set; }

        public Item(int itemSequel, string itemCode, string description, double quantity, string taxClassification, string icmsTax, double unitPrice, string ipiTax, string deliverTime)
        {
            ItemSequel = itemSequel;
            ItemCode = itemCode;
            Description = description;
            Quantity = quantity;
            TaxClassification = taxClassification;
            IcmsTax = icmsTax;
            UnitPrice = unitPrice;
            IpiTax = ipiTax;
            DeliverTime = deliverTime;
        }

        public double GetTotalPrice()
        {
            return Quantity * UnitPrice;
        }
    }
}
