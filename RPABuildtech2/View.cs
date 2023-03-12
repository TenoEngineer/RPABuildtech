using System;
using System.Windows.Forms;


namespace RPABuildtech2
{
    public partial class View : Form
    {
        public View()
        {
            InitializeComponent();
        }

        private void ButtonCancel(object sender, EventArgs e)
        {
            Close();
        }

        private void ButtonOK(object sender, EventArgs e)
        {
            Commands.RunReport(textBox1.Text, textBox2.Text, checkBox1.Checked, checkBox2.Checked);
            Close();
        }

        private void View_Load(object sender, EventArgs e)
        {

        }

    }
}
