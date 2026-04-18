using System.Drawing;
using System.Windows.Forms;

namespace NmapInventory
{
    public class InputDialog : Form
    {
        public string Value1 { get; set; }
        public string Value2 { get; set; }

        public InputDialog(string title, string label1, string label2)
        {
            Text = title;
            Width = 450; Height = 220;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;

            var txt1 = new TextBox { Location = new Point(20, 45), Width = 400 };
            var txt2 = new TextBox { Location = new Point(20, 105), Width = 400, Multiline = true, Height = 50 };
            var okBtn = new Button { Text = "OK", Location = new Point(220, 165), Width = 90, Height = 28, DialogResult = DialogResult.OK };
            var cancelBtn = new Button { Text = "Abbrechen", Location = new Point(320, 165), Width = 100, Height = 28, DialogResult = DialogResult.Cancel };

            okBtn.Click += (s, e) => { Value1 = txt1.Text; Value2 = txt2.Text; };
            this.Load += (s, e) => { txt1.Text = Value1 ?? ""; txt2.Text = Value2 ?? ""; };

            Controls.AddRange(new Control[] {
                new Label { Text = label1, Location = new Point(20, 20), AutoSize = true },
                txt1,
                new Label { Text = label2, Location = new Point(20, 80), AutoSize = true },
                txt2,
                okBtn, cancelBtn
            });
            AcceptButton = okBtn; CancelButton = cancelBtn;
        }
    }
}
