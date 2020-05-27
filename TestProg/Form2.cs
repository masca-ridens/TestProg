using System;
using System.Linq;
using System.Windows.Forms;

namespace TestProg
{
    public partial class Form2 : Form
    {
        public string t;
        public Form2()
        {
            InitializeComponent();
        }

        private void ButtonOk_Click(object sender, EventArgs e)
        {
            var checkedButton = groupBox1.Controls.OfType<RadioButton>().FirstOrDefault(r => r.Checked);
            t = checkedButton.Text;
        }
    }
}
