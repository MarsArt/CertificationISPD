using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CertificationISPD
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            KYZ = new Form2();
            OZ = new Form3();
            OAY = new Form4();
            PI = new ProgrammInspect();
        }
        public Form2 KYZ;
        public Form3 OZ;
        public Form4 OAY;
        public ProgrammInspect PI;
        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            KYZ.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OZ.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OAY.fillzero();
            OAY.ShowDialog();
            OAY.fillzero();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            PI.ShowDialog();
        }
    }
}
