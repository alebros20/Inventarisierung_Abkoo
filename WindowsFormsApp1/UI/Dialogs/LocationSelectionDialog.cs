using System.Drawing;
using System.Windows.Forms;

namespace NmapInventory
{
    public class LocationSelectionDialog : Form
    {
        public int SelectedLocationID { get; private set; }
        private DatabaseManager dbManager;
        private int sourceLocationID;
        private TreeView locationTree;

        public LocationSelectionDialog(DatabaseManager dbManager, int sourceLocationID)
        {
            this.dbManager = dbManager;
            this.sourceLocationID = sourceLocationID;
            Text = "Ziel-Location auswählen";
            Width = 400; Height = 500;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            Font = new Font("Segoe UI", 10);

            var label = new Label { Text = "Ziel-Abteilung auswählen:", Location = new Point(10, 10), AutoSize = true, Font = new Font("Segoe UI", 11, FontStyle.Bold) };
            locationTree = new TreeView { Location = new Point(10, 40), Width = 360, Height = 380, Font = new Font("Segoe UI", 10) };
            PopulateLocationTree();

            var okBtn     = new Button { Text = "Verschieben", Location = new Point(140, 430), Width = 110, Height = 28, DialogResult = DialogResult.OK, Font = new Font("Segoe UI", 10) };
            var cancelBtn = new Button { Text = "Abbrechen",   Location = new Point(260, 430), Width = 110, Height = 28, DialogResult = DialogResult.Cancel, Font = new Font("Segoe UI", 10) };
            okBtn.Click += (s, e) =>
            {
                var nodeData = locationTree.SelectedNode?.Tag as NodeData;
                if (nodeData?.Type == "Location") SelectedLocationID = nodeData.ID;
                else { MessageBox.Show("Bitte wähle eine Abteilung aus!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information); DialogResult = DialogResult.None; }
            };
            Controls.AddRange(new Control[] { label, locationTree, okBtn, cancelBtn });
            AcceptButton = okBtn; CancelButton = cancelBtn;
        }

        private void PopulateLocationTree()
        {
            locationTree.Nodes.Clear();
            foreach (var customer in dbManager.GetCustomers())
            {
                var customerNode = new TreeNode($"👤 {customer.Name}") { Tag = new NodeData { Type = "Customer", ID = customer.ID } };
                var locations = dbManager.GetLocationsByCustomer(customer.ID);
                for (int i = 0; i < locations.Count; i++)
                    AddLocationNode(customerNode, locations[i], $"S{i + 1}");
                locationTree.Nodes.Add(customerNode);
            }
            locationTree.ExpandAll();
        }

        private void AddLocationNode(TreeNode parent, Location location, string shortID)
        {
            var children = dbManager.GetChildLocations(location.ID);
            if (location.ID == sourceLocationID)
            {
                for (int i = 0; i < children.Count; i++)
                {
                    string cid = location.Level == 0 ? $"{shortID}.A{i + 1}" : $"{shortID}.{i + 1}";
                    AddLocationNode(parent, children[i], cid);
                }
                return;
            }
            string icon = location.Level == 0 ? "🏢" : "📂";
            var node = new TreeNode($"[{shortID}] {icon} {location.Name}") { Tag = new NodeData { Type = "Location", ID = location.ID } };
            for (int i = 0; i < children.Count; i++)
            {
                string cid = location.Level == 0 ? $"{shortID}.A{i + 1}" : $"{shortID}.{i + 1}";
                AddLocationNode(node, children[i], cid);
            }
            parent.Nodes.Add(node);
        }
    }
}
