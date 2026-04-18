using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace NmapInventory
{
    public class CustomerSelectionForm : Form
    {
        public int SelectedCustomerID { get; private set; }

        public CustomerSelectionForm(List<Customer> customers)
        {
            Text = "Kunde auswählen";
            Width = 400; Height = 200;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;

            var label = new Label { Text = "Wähle einen Kunden aus:", Location = new Point(20, 20), AutoSize = true, Font = new Font("Segoe UI", 11) };
            var combo = new ComboBox { Location = new Point(20, 50), Width = 350, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 10) };
            foreach (var customer in customers)
                combo.Items.Add(new ComboCustomerItem { ID = customer.ID, Name = customer.Name });
            if (combo.Items.Count > 0) combo.SelectedIndex = 0;

            var okBtn = new Button { Text = "OK", Location = new Point(150, 120), Width = 100, Height = 28, DialogResult = DialogResult.OK, Font = new Font("Segoe UI", 10) };
            var cancelBtn = new Button { Text = "Abbrechen", Location = new Point(260, 120), Width = 100, Height = 28, DialogResult = DialogResult.Cancel, Font = new Font("Segoe UI", 10) };

            okBtn.Click += (s, e) =>
            {
                if (combo.SelectedItem is ComboCustomerItem item)
                    SelectedCustomerID = item.ID;
            };

            Controls.AddRange(new Control[] { label, combo, okBtn, cancelBtn });
            AcceptButton = okBtn; CancelButton = cancelBtn;
        }

        // Typsicheres ComboBox-Item (kein dynamic mehr)
        private class ComboCustomerItem
        {
            public int ID { get; set; }
            public string Name { get; set; }
            public override string ToString() => Name;
        }
    }
}
