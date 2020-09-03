using System;
using System.Linq;
using System.Windows.Forms;

namespace SharedMailboxes
{
    public partial class Form1 : Form
    {
        private String Script = "";
        public Form1()
        {
            InitializeComponent();
            Script = "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private String generateString()
        {
            /*switch (true)
            {
                default:
            }*/
            return Script;
        }
    }
}
