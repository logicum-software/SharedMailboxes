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
        private String Script1e = " -OrganizationalUnit ";
        private String Script1f = "'caritas-brilon-wfb.de/03_User/74_Raumpostfach_Raumkalender'";
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
                Script1f = "'caritas-brilon-wfb.de/03_User/75_Geraetepostfaecher'";
                Script1h = "MBDB_EX2014_A-K' ";
                Script1i = "-Equipment";
                textBox2.Text = generateString();
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                Script1b = "Raum ";
                Script1f = "'caritas-brilon-wfb.de/03_User/74_Raumpostfach_Raumkalender'";
                Script1h = "MBDB_EX2014_L-Z' ";
                Script1i = "-Room";
                textBox2.Text = generateString();
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                Script1b = "Kalender ";
                Script1f = "'caritas-brilon-wfb.de/03_User/75_Geraetepostfaecher'";
                Script1h = "MBDB_EX2014_A-K' ";
                Script1i = "-Equipment";
                textBox2.Text = generateString();
            }
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton4.Checked)
            {
                Script1c = "WfbM-MW_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton5.Checked)
            {
                Script1c = "WfbM-HG_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton6.Checked)
            {
                Script1c = "WfbM-IDL1_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton7.Checked)
            {
                Script1c = "WfbM-IDL2_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton8.Checked)
            {
                Script1c = "WfbM-MSB_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton9_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton9.Checked)
            {
                Script1c = "WfbM-WTB_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton10_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton10.Checked)
            {
                Script1c = "WH-HH_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton11_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton11.Checked)
            {
                Script1c = "WH-EH_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton12_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton12.Checked)
            {
                Script1c = "WH-NH_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton13_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton13.Checked)
            {
                Script1c = "WH-DEH_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton14_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton14.Checked)
            {
                Script1c = "WH-LH_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton15_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton15.Checked)
            {
                Script1c = "WH-HN_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton16_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton16.Checked)
            {
                Script1c = "AKH-SOZ-BRI_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton17_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton17.Checked)
            {
                Script1c = "AKH-SOZ-MSB_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton18_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton18.Checked)
            {
                Script1c = "AKH-SOZ-OLB_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton19_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton19.Checked)
            {
                Script1c = "AKH-SOZ-WTB_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton20_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton20.Checked)
            {
                Script1c = "AKH-SOZ-HLB_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton21_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton21.Checked)
            {
                Script1c = "AKH-SOZ-MDB_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton25_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton25.Checked)
            {
                Script1c = "AKH-HNR_";
                textBox2.Text = generateString();
            }

        }

        private void radioButton22_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton22.Checked)
            {
                Script1c = "AKH-TPH_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton23_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton23.Checked)
            {
                Script1c = "AKH-STE_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton24_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton24.Checked)
            {
                Script1c = "AKH-STJ_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton26_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton26.Checked)
            {
                Script1c = "MKK-WTB_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton27_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton27.Checked)
            {
                Script1c = "MKK-BW_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton28_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton28.Checked)
            {
                Script1c = "BH-ABW-BRI_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton29_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton29.Checked)
            {
                Script1c = "BH-ABW-WTB_";
                textBox2.Text = generateString();
            }
        }

        private void radioButton30_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton30.Checked)
            {
                Script1c = "GS_";
                textBox2.Text = generateString();
            }
        }
    }
}