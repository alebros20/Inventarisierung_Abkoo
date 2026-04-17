using System;
using System.Collections.Generic;
using System.Linq;

namespace NmapInventory
{
    public class DataExporter
    {
        public void Export(string filePath, List<DeviceInfo> devices, List<SoftwareInfo> software, string hardware)
        {
            if (filePath.EndsWith(".json"))
            {
                var data = new { Zeitstempel = DateTime.Now, Geräte = devices, Software = software, Hardware = hardware };
                System.IO.File.WriteAllText(filePath, Newtonsoft.Json.JsonConvert.SerializeObject(data, Newtonsoft.Json.Formatting.Indented));
            }
            else
            {
                var doc = new System.Xml.Linq.XDocument(
                    new System.Xml.Linq.XDeclaration("1.0", "utf-8", "yes"),
                    new System.Xml.Linq.XElement("Inventar",
                        new System.Xml.Linq.XElement("Zeitstempel", DateTime.Now),
                        new System.Xml.Linq.XElement("Geräte", devices.Select(d => new System.Xml.Linq.XElement("Gerät",
                            new System.Xml.Linq.XElement("IP", d.IP),
                            new System.Xml.Linq.XElement("Hostname", d.Hostname ?? "")))),
                        new System.Xml.Linq.XElement("Hardware", new System.Xml.Linq.XCData(hardware))));
                doc.Save(filePath);
            }
        }
    }
}
