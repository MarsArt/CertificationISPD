using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CertificationISPD
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox8_Enter(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                WordReportGenerator WR = new WordReportGenerator();
                if (saveFileDialog1.ShowDialog() == DialogResult.OK || saveFileDialog1.FileName == "")
                {
                    WR.PathToSave = saveFileDialog1.FileName+".docx";
                    if (File.Exists("NewAct.docx"))
                    {
                        WR.template = "NewAct.docx";
                        string kategor = (radioButton1.Checked) ? radioButton1.Text :
                                     (radioButton2.Checked) ? radioButton2.Text :
                                     (radioButton3.Checked) ? radioButton3.Text :
                                     (radioButton4.Checked) ? radioButton4.Text : "";
                        WR.Add("@kategor", kategor);
                        string obem = (radioButton5.Checked) ? radioButton5.Text :
                                      (radioButton6.Checked) ? radioButton6.Text : "";
                        WR.Add("@obem", obem);

                        string tipsub = (radioButton7.Checked) ? radioButton7.Text :
                                       (radioButton8.Checked) ? radioButton8.Text : "";
                        WR.Add("@tipsub", tipsub);

                        string seti = (radioButton9.Checked) ? radioButton9.Text :
                                       (radioButton10.Checked) ? radioButton10.Text : "";
                        WR.Add("@seti", seti);

                        string structure = (radioButton11.Checked) ? radioButton11.Text :
                                           (radioButton12.Checked) ? radioButton12.Text :
                                       (radioButton13.Checked) ? radioButton13.Text : "";
                        WR.Add("@structure", structure);

                        string rejobrabot = (radioButton14.Checked) ? radioButton14.Text :
                                       (radioButton15.Checked) ? radioButton15.Text : "";
                        WR.Add("@rejobrabot", rejobrabot);

                        string dostyp = (radioButton16.Checked) ? radioButton16.Text :
                                       (radioButton17.Checked) ? radioButton17.Text : "";
                        WR.Add("@dostyp", dostyp);

                        string mesto = (radioButton18.Checked) ? radioButton18.Text :
                                       (radioButton19.Checked) ? radioButton19.Text : "";
                        WR.Add("@mesto", mesto);

                        string tipygr = (radioButton20.Checked) ? radioButton20.Text :
                                        (radioButton21.Checked) ? radioButton21.Text :
                                       (radioButton22.Checked) ? radioButton22.Text : "";
                        WR.Add("@tipygr", tipygr);
                        WR.Add("@UrovenZash", textBox1.Text);
                        string harakter = "";
                        if (checkBox1.Checked)
                            harakter = (harakter == "") ? checkBox1.Text : harakter += ", " + checkBox1.Text;
                        if (checkBox2.Checked)
                            harakter = (harakter == "") ? checkBox2.Text : harakter += ", " + checkBox2.Text;
                        if (checkBox3.Checked)
                            harakter = (harakter == "") ? checkBox3.Text : harakter += ", " + checkBox3.Text;
                        WR.Add("@harakter", harakter);

                        WR.Generate();
                        MessageBox.Show("Отчет успешно создан! Путь документа: " +WR.PathToSave, "Отчет");
                    }
                }
                else MessageBox.Show("Вы не указали путь для отчета! Отчет не был создан", "Ошибка!!!");

            }
            catch (Exception ex)
            { MessageBox.Show(ex.Message); }

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string UZ = "";
            if ((radioButton1.Checked && radioButton8.Checked && radioButton6.Checked) && (radioButton20.Checked || radioButton21.Checked)) UZ = "1 УЗ";
            if (radioButton1.Checked && radioButton8.Checked && radioButton6.Checked && radioButton22.Checked) UZ = "2 УЗ";
            if (radioButton1.Checked && radioButton8.Checked && radioButton5.Checked && radioButton20.Checked) UZ = "1 УЗ";
            if (radioButton1.Checked && radioButton7.Checked && radioButton20.Checked) UZ = "1 УЗ";
            if (radioButton1.Checked && radioButton7.Checked && radioButton21.Checked) UZ = "2 УЗ";
            if (radioButton1.Checked && radioButton7.Checked && radioButton22.Checked) UZ = "3 УЗ";
            if (radioButton2.Checked && radioButton20.Checked) UZ = "1 УЗ";
            if (radioButton2.Checked && radioButton21.Checked) UZ = "2 УЗ";
            if (radioButton2.Checked && radioButton22.Checked) UZ = "3 УЗ";
            if (radioButton4.Checked && radioButton8.Checked && radioButton6.Checked && radioButton20.Checked) UZ = "1 УЗ";
            if (radioButton4.Checked && radioButton8.Checked && radioButton6.Checked && radioButton21.Checked) UZ = "2 УЗ";
            if (radioButton4.Checked && radioButton8.Checked && radioButton6.Checked && radioButton22.Checked) UZ = "3 УЗ";
            if (radioButton4.Checked && radioButton8.Checked && radioButton5.Checked && radioButton20.Checked) UZ = "2 УЗ";
            if (radioButton4.Checked && radioButton8.Checked && radioButton5.Checked && radioButton21.Checked) UZ = "3 УЗ";
            if (radioButton4.Checked && radioButton8.Checked && radioButton5.Checked && radioButton22.Checked) UZ = "4 УЗ";
            if (radioButton4.Checked && radioButton7.Checked && radioButton20.Checked) UZ = "2 УЗ";
            if (radioButton4.Checked && radioButton7.Checked && radioButton21.Checked) UZ = "3 УЗ";
            if (radioButton4.Checked && radioButton7.Checked && radioButton22.Checked) UZ = "4 УЗ";
            if ((radioButton3.Checked && radioButton8.Checked && radioButton6.Checked) && (radioButton20.Checked || radioButton21.Checked)) UZ = "2 УЗ";
            if (radioButton3.Checked && radioButton8.Checked && radioButton6.Checked && radioButton22.Checked) UZ = "4 УЗ";
            if (radioButton3.Checked && radioButton8.Checked && radioButton5.Checked && radioButton20.Checked) UZ = "2 УЗ";
            if (radioButton3.Checked && radioButton8.Checked && radioButton5.Checked && radioButton22.Checked) UZ = "4 УЗ";
            if (radioButton3.Checked && radioButton7.Checked && radioButton20.Checked) UZ = "2 УЗ";
            if (radioButton3.Checked && radioButton7.Checked && radioButton21.Checked) UZ = "3 УЗ";
            if (radioButton3.Checked && radioButton7.Checked && radioButton22.Checked) UZ = "4 УЗ";
            textBox1.Text = UZ;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
