namespace NmapInventory
{
    public class Customer
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
    }

    public class Location
    {
        public int ID { get; set; }
        public int CustomerID { get; set; }
        public int ParentID { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
        public int Level { get; set; }
    }

    public class LocationIP
    {
        public int ID { get; set; }
        public int LocationID { get; set; }
        public string IPAddress { get; set; }
        public string WorkstationName { get; set; }
    }

    public class NodeData
    {
        public string Type { get; set; }
        public int ID { get; set; }
        public object Data { get; set; }
    }
}
