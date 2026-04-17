namespace NmapInventory
{
    public class ComboItem
    {
        public int ID { get; set; }
        public string Display { get; set; }
        public string FilePath { get; set; }
        public override string ToString() => Display;
    }
}
