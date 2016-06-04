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
    public partial class SetSourceUpdate : Form
    {
        public SetSourceUpdate()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void SetSourceUpdate_FormClosing(object sender, FormClosingEventArgs e)
        {
            ProgrammInspect F1 = new ProgrammInspect();
            F1.loader.saveConfig();
        }
    }
}
