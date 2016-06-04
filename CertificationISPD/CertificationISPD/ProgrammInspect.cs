using Microsoft.Win32;
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
    public partial class ProgrammInspect : Form
    {
        public Loader loader;
        public WindowsReaderProgramSearcher WRPS;

        public List<SetupProgramDescr> SetupProgramCertificateList;
        
        public ProgrammInspect()
        {
            InitializeComponent();
            loader = new Loader();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.source != "")
                loader.loadConfig();
        }
      /*  private void button1_Click(object sender, EventArgs e)
        {
            loader.addSourceItem("https://docs.google.com/spreadsheets/d/12AQgCDP47NME0NxpByY2lP0AeyL6BOtTV9VmWy09Apw/export?format=xlsx&id=12AQgCDP47NME0NxpByY2lP0AeyL6BOtTV9VmWy09Apw", "FSTEC.xlsx");
            loader.saveConfig();

        }
        private void button2_Click(object sender, EventArgs e)
        {
            loader.loadConfig();
        }*/
        private void обновитьБазуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            loader.LoadFiles();
        }
        private void считатьСписокУстановленыхПрограммToolStripMenuItem_Click(object sender, EventArgs e)
        {
            WRPS = new WindowsReaderProgramSearcher();
            SetupProgramCertificateList = new List<SetupProgramDescr>();
            WRPS.ReadProgram();
            WRPS.fillListCert();
            setupProgramDescrBindingSource.DataSource = WRPS.ListSetupWidthRelation;
            foreach (SetupProgramDescr SPD in WRPS.ListSetupWidthRelation) {
                if (SPD.RelationProgram.Count == 1) {
                    SetupProgramCertificateList.Add(SPD);
                }
            }
            setupProgramDescrBindingSource1.DataSource = SetupProgramCertificateList;
        }   
        private void отобразитьСертифицированыеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FSTECFindCertificationProgram CERRider = new FSTECFindCertificationProgram();
            CERRider.getListCert("Windows NT 4.0");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SetupProgramDescr SetupProgramDescrCert = new SetupProgramDescr((SetupProgramDescr)setupProgramDescrBindingSource.Current, (CertifProgramDescr)relationProgramBindingSource.Current);
            if (!setupProgramDescrBindingSource1.Contains(SetupProgramDescrCert))
                    setupProgramDescrBindingSource1.Add(SetupProgramDescrCert);
            else MessageBox.Show("Такой элемент существует");
            setupProgramDescrBindingSource1.Sort = "Name ASC";
            dataGridView1.DataSource = setupProgramDescrBindingSource1;


        }

        private void setupProgramDescrBindingSource1_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            setupProgramDescrBindingSource1.RemoveCurrent();
            
        }
    }
}
