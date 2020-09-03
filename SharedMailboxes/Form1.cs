using System;
using System.Linq;
using System.Windows.Forms;

namespace SharedMailboxes
{
    public partial class Form1 : Form
    {
        private String Script1a = "New-Mailbox -Name '";
        private String Script1b = "Raum ";
        private String Script1c = "WfbM-MW_";
        private String Script1d = " Alias '";
        private String Script1e = " -OrganizationalUnit '";
        private String Script1f = "caritas-brilon-wfb.de/03_User/74_Raumpostfach_Raumkalender'";
        private String Script1g = " -UserPrincipalName '";
        private String Script1h = "MBDB_EX2014_L-Z' ";
        private String Script1i = "-Room";

        public Form1()
        {
            InitializeComponent();
            textBox2.Text = generateString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private String generateString()
        {
            return Script1a + Script1b + Script1c + textBox1.Text  + "'" + Script1d + Script1c + textBox1.Text +
                "'" + Script1e + Script1f + Script1g + Script1c + textBox1.Text + "@caritas-brilon-wfb.de'" +
                "-SamAccountName '" + Script1c + textBox1.Text + "' -FirstName '' -Initials '' -LastName '' " +
                "-Database '" + Script1h + Script1i;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //Script1e = Script1b + label2.Text + "'";
            textBox2.Text = generateString();
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked)
            {
                Script1b = "Fzg-";
                textBox2.Text = generateString();
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                Script1b = "Raum ";
                textBox2.Text = generateString();
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                Script1b = "Kalender ";
                textBox2.Text = generateString();
            }
        }
    }
}